using EnvDTE;
using EnvDTE90a;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Tagging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Stopwatch = System.Diagnostics.Stopwatch;

namespace InlineCppVarDbg
{
    internal sealed class InlineValueTagger : ITagger<IntraTextAdornmentTag>
    {
        private const string DefaultChipBackgroundHex = "#E5E5E5";
        private const string DefaultUninitializedChipBackgroundHex = "#FFF59D";
        private const string DefaultChangedAccentHex = "#FF8FB1";
        private const int MaxManualFadeSteps = 5;
        private const int DefaultPerExpressionTimeoutMs = 10;
        private const long PerfLogThresholdMs = 250;
        private const long RequestedProfileLogThresholdMs = 3000;
        private const long SlowEvaluationThresholdMs = 25;
        private const string NullValueMarker = "\u0001N:";
        private const string UninitializedValueMarker = "\u0001U:";
        private static readonly Guid PerfOutputPaneGuid = new Guid("2A09B390-E700-42DE-9A4D-DC37F74B71F2");

        private static readonly HashSet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".c", ".cc", ".cpp", ".cxx", ".h", ".hh", ".hpp", ".hxx", ".inl"
        };

        private static readonly HashSet<string> PrimitiveTypeTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "bool", "char", "char8_t", "char16_t", "char32_t", "wchar_t", "short", "int", "long",
            "float", "double", "void", "__int8", "__int16", "__int32", "__int64",
            "int8_t", "int16_t", "int32_t", "int64_t",
            "uint8", "uint16", "uint32", "uint64",
            "uint8_t", "uint16_t", "uint32_t", "uint64_t",
            "size_t", "ptrdiff_t"
        };

        private static readonly HashSet<string> TypeNoiseTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "const", "volatile", "signed", "unsigned", "class", "struct", "union", "enum",
            "__ptr64", "__restrict", "__unaligned", "__cdecl", "__stdcall", "__fastcall", "__thiscall"
        };

        private static readonly HashSet<string> NonFunctionBlockKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "if", "for", "while", "switch", "catch"
        };

        private static readonly Color DefaultChipBackgroundColor = ParseColorOrFallback(DefaultChipBackgroundHex, Colors.LightGray);
        private static readonly Color DefaultUninitializedChipBackgroundColor = ParseColorOrFallback(DefaultUninitializedChipBackgroundHex, Colors.LightYellow);
        private static readonly Color DefaultChangedAccentColor = ParseColorOrFallback(DefaultChangedAccentHex, Colors.Red);
        private static readonly Brush DefaultChipBackground = CreateFrozenBrush(DefaultChipBackgroundColor);
        private static readonly Brush DefaultUninitializedChipBackground = CreateFrozenBrush(DefaultUninitializedChipBackgroundColor);
        private static readonly Brush ChipBorder = CreateFrozenBrush("#C8C8C8");
        private static readonly Brush ValueForeground = CreateFrozenBrush("#0B5CAD");
        private static readonly Brush NullValueForeground = CreateFrozenBrush("#C62828");
        private static readonly Brush StatusForeground = CreateFrozenBrush("#6A6A6A");

        private readonly IWpfTextView textView;
        private readonly ITextBuffer textBuffer;
        private readonly string documentPath;
        private readonly DebuggerBridge debuggerBridge;
        private readonly InlineValuesSettings settings;
        private readonly IClassifier classifier;
        private readonly string normalizedDocumentPath;
        private Brush chipBackground = DefaultChipBackground;
        private Brush uninitializedChipBackground = DefaultUninitializedChipBackground;
        private Color chipBackgroundColor = DefaultChipBackgroundColor;
        private Color uninitializedChipBackgroundColor = DefaultUninitializedChipBackgroundColor;
        private Color changedAccentColor = DefaultChangedAccentColor;
        private double chipFontSize = 10.0;

        private ITextSnapshot cachedSnapshot;
        private int cachedDebuggerVersion = -1;
        private ComputedTagData cachedData;
        private bool disposed;
        private int lastComparedDebuggerVersion = -1;
        private Dictionary<string, string> lastResolvedValues = new Dictionary<string, string>(StringComparer.Ordinal);
        private Dictionary<string, int> lastHighlightLevels = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly HashSet<string> manualGetterRequests = new HashSet<string>(StringComparer.Ordinal);
        [ThreadStatic]
        private static PerfSession currentPerfSession;

        public InlineValueTagger(
            IWpfTextView textView,
            ITextBuffer textBuffer,
            string documentPath,
            DebuggerBridge debuggerBridge,
            InlineValuesSettings settings,
            IClassifier classifier)
        {
            this.textView = textView;
            this.textBuffer = textBuffer;
            this.documentPath = documentPath;
            this.debuggerBridge = debuggerBridge;
            this.settings = settings;
            this.classifier = classifier;
            normalizedDocumentPath = NormalizePath(documentPath);

            textBuffer.Changed += OnBufferChanged;
            textView.Closed += OnTextViewClosed;
            textView.VisualElement.PreviewMouseLeftButtonDown += OnPreviewMouseLeftButtonDown;
            debuggerBridge.DebugStateChanged += OnDebugStateChanged;
            settings.SettingsChanged += OnSettingsChanged;
            UpdateAppearanceSettings();
        }

        public event EventHandler<SnapshotSpanEventArgs> TagsChanged;

        public IEnumerable<ITagSpan<IntraTextAdornmentTag>> GetTags(NormalizedSnapshotSpanCollection spans)
        {
            if (disposed || spans.Count == 0)
            {
                yield break;
            }

            if (!IsSupportedCppPath(documentPath))
            {
                yield break;
            }

            if (!textView.VisualElement.Dispatcher.CheckAccess())
            {
                yield break;
            }

            ThreadHelper.ThrowIfNotOnUIThread();
            ITextSnapshot snapshot = spans[0].Snapshot;
            ComputedTagData data = GetOrComputeData(snapshot);
            if (data == null || data.Tags.Count == 0)
            {
                yield break;
            }

            bool intersectsActiveLine = false;
            foreach (SnapshotSpan requested in spans)
            {
                if (requested.IntersectsWith(data.CoveredSpan))
                {
                    intersectsActiveLine = true;
                    break;
                }
            }

            if (!intersectsActiveLine)
            {
                yield break;
            }

            foreach (TagEntry entry in data.Tags)
            {
                int position = entry.Position;
                if (position < 0 || position > snapshot.Length)
                {
                    continue;
                }

                var pointSpan = new SnapshotSpan(snapshot, new Span(position, 0));
                var adornment = CreateAdornment(entry.DisplayText, entry.HighlightLevel, entry.DisplayKind);
                yield return new TagSpan<IntraTextAdornmentTag>(
                    pointSpan,
                    new IntraTextAdornmentTag(adornment, null, entry.Affinity));
            }
        }

        private static Brush CreateFrozenBrush(string hexColor)
        {
            var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hexColor));
            brush.Freeze();
            return brush;
        }

        private static Brush CreateFrozenBrush(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        private static bool IsSupportedCppPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            string extension = System.IO.Path.GetExtension(path);
            return SupportedExtensions.Contains(extension ?? string.Empty);
        }

        private ComputedTagData GetOrComputeData(ITextSnapshot snapshot)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            int version = debuggerBridge.Version;
            if (ReferenceEquals(cachedSnapshot, snapshot) && cachedDebuggerVersion == version && cachedData != null)
            {
                return cachedData;
            }

            cachedSnapshot = snapshot;
            cachedDebuggerVersion = version;
            cachedData = ComputeData(snapshot);
            return cachedData;
        }

        private ComputedTagData ComputeData(ITextSnapshot snapshot)
        {
            if (!settings.IsEnabled)
            {
                return null;
            }

            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (!debuggerBridge.TryGetCurrentBreakContext(out DebuggerBridge.BreakContext context))
            {
                if (debuggerBridge.IsInDesignMode())
                {
                    ClearChangeTracking();
                }

                return null;
            }

            if (!PathsEqual(normalizedDocumentPath, NormalizePath(context.FileName)))
            {
                return null;
            }

            int lineNumber = context.LineNumber - 1;
            if (lineNumber < 0 || lineNumber >= snapshot.LineCount)
            {
                return null;
            }

            bool manualVariableFunctionSweep = debuggerBridge.IsManualVariableFunctionSweepRequested;
            EvaluationPolicy evaluationPolicy = CreateEvaluationPolicy(settings.PreviousLineCount, manualVariableFunctionSweep);
            int functionStartLineNumber = FindEnclosingFunctionStartLine(snapshot, lineNumber);
            bool manualGetterFunctionSweep = debuggerBridge.IsManualGetterFunctionSweepRequested;
            int functionEndLineNumber = functionStartLineNumber >= 0
                ? FindEnclosingFunctionEndLine(snapshot, functionStartLineNumber)
                : lineNumber;
            int previousLineCount = evaluationPolicy.PreviousLineCount;
            int functionScopeStartLine = functionStartLineNumber >= 0 ? functionStartLineNumber : 0;
            int regularStartLineNumber = manualVariableFunctionSweep
                ? functionScopeStartLine
                : Math.Max(functionScopeStartLine, lineNumber - previousLineCount);
            int regularEndLineNumber = manualVariableFunctionSweep
                ? Math.Max(lineNumber, functionEndLineNumber)
                : lineNumber;
            int getterStartLineNumber = manualGetterFunctionSweep ? functionScopeStartLine : regularStartLineNumber;
            int getterEndLineNumber = manualGetterFunctionSweep ? Math.Max(lineNumber, functionEndLineNumber) : regularEndLineNumber;
            int coveredStartLineNumber = Math.Min(regularStartLineNumber, getterStartLineNumber);
            int coveredEndLineNumber = Math.Max(regularEndLineNumber, getterEndLineNumber);
            HashSet<string> parameterNames = GetStackFrameParameterNames(context.StackFrame);
            InlineValueNumericDisplayMode numericDisplayMode = settings.NumericDisplayMode;
            InlineValueEvaluationKinds evaluationKinds = settings.EvaluationKinds;
            InlineValueRuleKinds ruleKinds = settings.RuleKinds;
            InlineValueTypeRuleKinds typeRuleKinds = settings.TypeRuleKinds;
            IReadOnlyList<InlineValueCustomRule> customRules = settings.CustomRules;
            DebuggerBridge.ProfileRequestKind profileRequestKind = debuggerBridge.TryConsumeProfileForNextEvaluation(out DebuggerBridge.ProfileRequestKind consumedProfileRequestKind)
                ? consumedProfileRequestKind
                : DebuggerBridge.ProfileRequestKind.None;

            var tokensByLine = new List<LineTokens>();
            var allTokens = new List<CppCurrentLineTokenizer.IdentifierToken>();
            var allGetterCalls = new List<GetterCallToken>();
            var allMemberAccesses = new List<GetterCallToken>();
            for (int lineIndex = coveredStartLineNumber; lineIndex <= coveredEndLineNumber; lineIndex++)
            {
                ITextSnapshotLine line = snapshot.GetLineFromLineNumber(lineIndex);
                if (!HasRuleKind(ruleKinds, InlineValueRuleKinds.ParseInactivePreprocessor) && IsInactivePreprocessorLine(line))
                {
                    continue;
                }

                string lineText = line.GetText();
                bool scanRegularTokens = lineIndex >= regularStartLineNumber && lineIndex <= regularEndLineNumber;
                bool scanGetterTokens = lineIndex >= getterStartLineNumber && lineIndex <= getterEndLineNumber;
                List<CppCurrentLineTokenizer.IdentifierToken> rawTokens =
                    (scanRegularTokens || scanGetterTokens)
                    ? CppCurrentLineTokenizer.TokenizeIdentifiers(lineText)
                    : new List<CppCurrentLineTokenizer.IdentifierToken>();
                var tokens = new List<CppCurrentLineTokenizer.IdentifierToken>(scanRegularTokens ? rawTokens.Count : 0);
                var getterCalls = new List<GetterCallToken>();
                var memberAccesses = new List<GetterCallToken>();
                foreach (CppCurrentLineTokenizer.IdentifierToken token in rawTokens)
                {
                    if (scanGetterTokens &&
                        TryParseGetterCall(lineText, token, out GetterCallToken getterCall) &&
                        TryBindVisibleDirectReturnGetter(snapshot, getterCall, out GetterCallToken boundGetterCall))
                    {
                        bool allowAutomaticGetter = HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.GetterCalls);
                        bool allowManualGetter = manualGetterFunctionSweep || manualGetterRequests.Contains(boundGetterCall.ExpressionText);
                        if (allowAutomaticGetter || allowManualGetter)
                        {
                            getterCalls.Add(boundGetterCall);
                        }
                    }

                    if (scanRegularTokens &&
                        HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.IndexedExpressions) &&
                        TryParseArrayIndexAccess(lineText, token, out GetterCallToken indexedCall))
                    {
                        getterCalls.Add(indexedCall);
                        continue;
                    }

                    if (!scanRegularTokens)
                    {
                        continue;
                    }

                    if (TryParseMemberAccessExpression(lineText, token, out GetterCallToken memberAccess))
                    {
                        memberAccesses.Add(memberAccess);
                        continue;
                    }

                    if (HasMemberAccessOperatorBeforeToken(lineText, token.Start))
                    {
                        continue;
                    }

                    if (IsPointerReceiverForGetterCall(lineText, token))
                    {
                        continue;
                    }

                    if (IsLikelyFunctionInvocation(lineText, token))
                    {
                        continue;
                    }

                    tokens.Add(token);
                }

                tokensByLine.Add(new LineTokens(line, tokens, getterCalls, memberAccesses));
                allTokens.AddRange(tokens);
                allGetterCalls.AddRange(getterCalls);
                allMemberAccesses.AddRange(memberAccesses);
            }

            if (allTokens.Count == 0 && allGetterCalls.Count == 0 && allMemberAccesses.Count == 0 && parameterNames.Count == 0)
            {
                ITextSnapshotLine startLine = snapshot.GetLineFromLineNumber(coveredStartLineNumber);
                ITextSnapshotLine endLine = snapshot.GetLineFromLineNumber(coveredEndLineNumber);
                int spanStart = startLine.Start.Position;
                int spanEnd = endLine.End.Position;
                return new ComputedTagData(new SnapshotSpan(snapshot, Span.FromBounds(spanStart, spanEnd)), new List<TagEntry>());
            }

            Dictionary<string, string> valueMap;
            Dictionary<string, string> getterValueMap;
            Dictionary<string, string> memberAccessValueMap;
            using (PerfSession perfSession = PerfSession.Start(documentPath, lineNumber + 1, debuggerBridge.Version, evaluationPolicy, profileRequestKind))
            {
                currentPerfSession = perfSession;
                try
                {
                    valueMap = ResolveValues(context.Debugger, context.StackFrame, allTokens, parameterNames, numericDisplayMode, evaluationKinds, ruleKinds, typeRuleKinds, customRules);
                    getterValueMap = allGetterCalls.Count > 0
                        ? ResolveGetterValues(context.Debugger, allGetterCalls, numericDisplayMode, evaluationKinds, ruleKinds, typeRuleKinds, customRules, manualGetterFunctionSweep, manualGetterRequests)
                        : new Dictionary<string, string>(StringComparer.Ordinal);
                    memberAccessValueMap = allMemberAccesses.Count > 0
                        ? ResolveExpressionValues(context.Debugger, allMemberAccesses, numericDisplayMode, evaluationKinds, ruleKinds, typeRuleKinds, customRules)
                        : new Dictionary<string, string>(StringComparer.Ordinal);
                }
                finally
                {
                    currentPerfSession = null;
                }
            }

            foreach (KeyValuePair<string, string> pair in getterValueMap)
            {
                if (!valueMap.ContainsKey(pair.Key))
                {
                    valueMap[pair.Key] = pair.Value;
                }
            }

            foreach (KeyValuePair<string, string> pair in memberAccessValueMap)
            {
                if (!valueMap.ContainsKey(pair.Key))
                {
                    valueMap[pair.Key] = pair.Value;
                }
            }

            Dictionary<string, int> highlightLevels = GetHighlightLevels(valueMap, debuggerBridge.Version);
            var firstEntryByIdentifier = new Dictionary<string, TagEntry>(StringComparer.Ordinal);
            var latestEntryByIdentifier = new Dictionary<string, TagEntry>(StringComparer.Ordinal);
            InlineValueDisplayMode displayMode = settings.DisplayMode;
            foreach (LineTokens lineTokens in tokensByLine)
            {
                var seenOnLine = new HashSet<string>(StringComparer.Ordinal);
                foreach (CppCurrentLineTokenizer.IdentifierToken token in lineTokens.Tokens)
                {
                    if (displayMode == InlineValueDisplayMode.EndOfLine && !seenOnLine.Add(token.Name))
                    {
                        continue;
                    }

                    if (!valueMap.TryGetValue(token.Name, out string value))
                    {
                        continue;
                    }

                    ValueDisplayKind displayKind = GetDisplayKind(value);
                    string cleanedValue = StripDisplayMarker(value);

                    int highlightLevel = 0;
                    if (highlightLevels.TryGetValue(token.Name, out int level))
                    {
                        highlightLevel = level;
                    }

                    int insertion;
                    PositionAffinity affinity;
                    string displayText;
                    if (displayMode == InlineValueDisplayMode.EndOfLine)
                    {
                        insertion = lineTokens.Line.End.Position;
                        affinity = PositionAffinity.Predecessor;
                        displayText = token.Name + ": " + cleanedValue;
                    }
                    else
                    {
                        insertion = lineTokens.Line.Start.Position + token.Start + token.Length;
                        affinity = PositionAffinity.Successor;
                        displayText = cleanedValue;
                    }

                    var tagEntry = new TagEntry(insertion, displayText, highlightLevel, affinity, displayKind);
                    if (!firstEntryByIdentifier.ContainsKey(token.Name))
                    {
                        // Use first visible appearance as the declaration/definition anchor.
                        firstEntryByIdentifier[token.Name] = tagEntry;
                    }

                    // Keep the latest visible appearance.
                    latestEntryByIdentifier[token.Name] = tagEntry;
                }

                var seenGetterOnLine = new HashSet<string>(StringComparer.Ordinal);
                foreach (GetterCallToken getterCall in lineTokens.GetterCalls)
                {
                    if (displayMode == InlineValueDisplayMode.EndOfLine && !seenGetterOnLine.Add(getterCall.ExpressionText))
                    {
                        continue;
                    }

                    if (!valueMap.TryGetValue(getterCall.ExpressionText, out string getterValue))
                    {
                        continue;
                    }

                    ValueDisplayKind displayKind = GetDisplayKind(getterValue);
                    string cleanedGetterValue = StripDisplayMarker(getterValue);

                    int highlightLevel = 0;
                    if (highlightLevels.TryGetValue(getterCall.ExpressionText, out int level))
                    {
                        highlightLevel = level;
                    }

                    int insertion;
                    PositionAffinity affinity;
                    string displayText;
                    if (displayMode == InlineValueDisplayMode.EndOfLine)
                    {
                        insertion = lineTokens.Line.End.Position;
                        affinity = PositionAffinity.Predecessor;
                        displayText = getterCall.DisplayLabel + ": " + cleanedGetterValue;
                    }
                    else
                    {
                        insertion = lineTokens.Line.Start.Position + getterCall.CallEnd;
                        affinity = PositionAffinity.Successor;
                        displayText = cleanedGetterValue;
                    }

                    var tagEntry = new TagEntry(insertion, displayText, highlightLevel, affinity, displayKind);
                    if (!firstEntryByIdentifier.ContainsKey(getterCall.ExpressionText))
                    {
                        firstEntryByIdentifier[getterCall.ExpressionText] = tagEntry;
                    }

                    latestEntryByIdentifier[getterCall.ExpressionText] = tagEntry;
                }

                var seenMemberAccessOnLine = new HashSet<string>(StringComparer.Ordinal);
                foreach (GetterCallToken memberAccess in lineTokens.MemberAccesses)
                {
                    if (displayMode == InlineValueDisplayMode.EndOfLine && !seenMemberAccessOnLine.Add(memberAccess.ExpressionText))
                    {
                        continue;
                    }

                    if (!valueMap.TryGetValue(memberAccess.ExpressionText, out string memberValue))
                    {
                        continue;
                    }

                    ValueDisplayKind displayKind = GetDisplayKind(memberValue);
                    string cleanedMemberValue = StripDisplayMarker(memberValue);

                    int highlightLevel = 0;
                    if (highlightLevels.TryGetValue(memberAccess.ExpressionText, out int level))
                    {
                        highlightLevel = level;
                    }

                    int insertion;
                    PositionAffinity affinity;
                    string displayText;
                    if (displayMode == InlineValueDisplayMode.EndOfLine)
                    {
                        insertion = lineTokens.Line.End.Position;
                        affinity = PositionAffinity.Predecessor;
                        displayText = memberAccess.DisplayLabel + ": " + cleanedMemberValue;
                    }
                    else
                    {
                        insertion = lineTokens.Line.Start.Position + memberAccess.CallEnd;
                        affinity = PositionAffinity.Successor;
                        displayText = cleanedMemberValue;
                    }

                    var tagEntry = new TagEntry(insertion, displayText, highlightLevel, affinity, displayKind);
                    if (!firstEntryByIdentifier.ContainsKey(memberAccess.ExpressionText))
                    {
                        firstEntryByIdentifier[memberAccess.ExpressionText] = tagEntry;
                    }

                    latestEntryByIdentifier[memberAccess.ExpressionText] = tagEntry;
                }
            }

            int parameterAnchorStartLine = int.MaxValue;
            if (HasRuleKind(ruleKinds, InlineValueRuleKinds.ParameterAnchors))
            {
                parameterAnchorStartLine = AddMissingParameterAnchorTags(
                    snapshot,
                    functionStartLineNumber,
                    parameterNames,
                    valueMap,
                    highlightLevels,
                    displayMode,
                    firstEntryByIdentifier,
                    latestEntryByIdentifier);
            }

            var tagEntries = new List<TagEntry>(latestEntryByIdentifier.Count * 2);
            foreach (KeyValuePair<string, TagEntry> pair in latestEntryByIdentifier)
            {
                tagEntries.Add(pair.Value);

                if (firstEntryByIdentifier.TryGetValue(pair.Key, out TagEntry firstEntry) &&
                    (firstEntry.Position != pair.Value.Position || firstEntry.Affinity != pair.Value.Affinity))
                {
                    tagEntries.Add(firstEntry);
                }
            }

            int finalCoveredStartLineNumber = coveredStartLineNumber;
            if (parameterAnchorStartLine != int.MaxValue)
            {
                finalCoveredStartLineNumber = Math.Min(finalCoveredStartLineNumber, parameterAnchorStartLine);
            }

            tagEntries = tagEntries
                .OrderBy(t => t.Position)
                .ThenBy(t => t.Affinity == PositionAffinity.Successor ? 1 : 0)
                .ToList();

            ITextSnapshotLine rangeStartLine = snapshot.GetLineFromLineNumber(finalCoveredStartLineNumber);
            ITextSnapshotLine rangeEndLine = snapshot.GetLineFromLineNumber(coveredEndLineNumber);
            int rangeStart = rangeStartLine.Start.Position;
            int rangeEnd = rangeEndLine.End.Position;

            return new ComputedTagData(new SnapshotSpan(snapshot, Span.FromBounds(rangeStart, rangeEnd)), tagEntries);
        }

        private bool IsInactivePreprocessorLine(ITextSnapshotLine line)
        {
            if (classifier == null || line == null || line.Length == 0)
            {
                return false;
            }

            SnapshotSpan lineSpan = line.Extent;
            IList<ClassificationSpan> spans;
            try
            {
                spans = classifier.GetClassificationSpans(lineSpan);
            }
            catch
            {
                return false;
            }

            foreach (ClassificationSpan classificationSpan in spans)
            {
                if (classificationSpan == null || classificationSpan.Span.Length == 0)
                {
                    continue;
                }

                if (!classificationSpan.Span.IntersectsWith(lineSpan))
                {
                    continue;
                }

                if (IsInactiveClassificationType(classificationSpan.ClassificationType))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsInactiveClassificationType(IClassificationType classificationType)
        {
            if (classificationType == null)
            {
                return false;
            }

            if (classificationType.IsOfType("excluded code") || classificationType.IsOfType("inactive code"))
            {
                return true;
            }

            string name = classificationType.Classification ?? string.Empty;
            if (name.IndexOf("excluded", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("inactive", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            foreach (IClassificationType baseType in classificationType.BaseTypes)
            {
                if (IsInactiveClassificationType(baseType))
                {
                    return true;
                }
            }

            return false;
        }

        private static Dictionary<string, string> ResolveValues(
            Debugger debugger,
            StackFrame2 stackFrame,
            List<CppCurrentLineTokenizer.IdentifierToken> tokens,
            HashSet<string> extraRequiredIdentifiers,
            InlineValueNumericDisplayMode numericDisplayMode,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds,
            IReadOnlyList<InlineValueCustomRule> customRules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            if (ShouldAbortEvaluation())
            {
                return values;
            }

            var requiredIdentifiers = new HashSet<string>(tokens.Select(t => t.Name), StringComparer.Ordinal);
            if (extraRequiredIdentifiers != null)
            {
                requiredIdentifiers.UnionWith(extraRequiredIdentifiers);
            }

            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                Expressions locals = stackFrame.Locals2[true];
                RecordPerf("StackFrame.Locals2", "Locals2[true]", stopwatch.ElapsedMilliseconds);
                AddExpressionValues(values, locals, requiredIdentifiers, 0, debugger, numericDisplayMode, EnumValueRenderMode.IntegerOnly, evaluationKinds, ruleKinds, typeRuleKinds, customRules);
            }
            catch (COMException)
            {
            }

            if (ShouldAbortEvaluation())
            {
                return values;
            }

            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                Expressions arguments = stackFrame.Arguments2[true];
                RecordPerf("StackFrame.Arguments2", "Arguments2[true]", stopwatch.ElapsedMilliseconds);
                AddExpressionValues(values, arguments, requiredIdentifiers, 0, debugger, numericDisplayMode, EnumValueRenderMode.IntegerOnly, evaluationKinds, ruleKinds, typeRuleKinds, customRules);
            }
            catch (COMException)
            {
            }

            if (!RuntimeAllowsFallbackExpressions() ||
                !HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.FallbackExpressions) ||
                ShouldAbortEvaluation())
            {
                return values;
            }

            foreach (string identifier in requiredIdentifiers)
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                if (values.ContainsKey(identifier))
                {
                    continue;
                }

                try
                {
                    EnvDTE.Expression expression = TimedGetExpression(debugger, identifier, 50, "Fallback.GetExpression");
                    if (expression != null && expression.IsValidValue)
                    {
                        TryAddExpressionValue(values, identifier, expression.Type, expression.Value, requiredIdentifiers, debugger, numericDisplayMode, EnumValueRenderMode.IntegerOnly, evaluationKinds, ruleKinds, typeRuleKinds, customRules);
                    }
                }
                catch (COMException)
                {
                }
            }

            return values;
        }

        private static Dictionary<string, string> ResolveGetterValues(
            Debugger debugger,
            List<GetterCallToken> getterCalls,
            InlineValueNumericDisplayMode numericDisplayMode,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds,
            IReadOnlyList<InlineValueCustomRule> customRules,
            bool allowManualGetterFunctionSweep,
            ISet<string> manualGetterExpressions)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            if (ShouldAbortEvaluation())
            {
                return values;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (GetterCallToken getterCall in getterCalls)
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                if (!seen.Add(getterCall.ExpressionText))
                {
                    continue;
                }

                bool isManualGetterRequest =
                    allowManualGetterFunctionSweep ||
                    (manualGetterExpressions != null && manualGetterExpressions.Contains(getterCall.ExpressionText));
                if (!isManualGetterRequest && !HasEvaluationKind(evaluationKinds, getterCall.EvaluationKind))
                {
                    continue;
                }

                try
                {
                    EnvDTE.Expression expression = TimedGetExpression(debugger, getterCall.EvaluationExpressionText, 50, "Getter.GetExpression");
                    if (expression == null || !expression.IsValidValue)
                    {
                        continue;
                    }

                    string normalized = NormalizeValue(expression.Value);
                    if (string.IsNullOrEmpty(normalized))
                    {
                        continue;
                    }

                    if (!IsDisplayableGetterReturnType(expression.Type, expression.Value, normalized))
                    {
                        continue;
                    }

                    if (!TryFormatEligibleValue(
                        debugger,
                        getterCall.EvaluationExpressionText,
                        expression.Type,
                        expression.Value,
                        normalized,
                        numericDisplayMode,
                        EnumValueRenderMode.IntegerOnly,
                        evaluationKinds,
                        ruleKinds,
                        typeRuleKinds,
                        customRules,
                        out string displayValue))
                    {
                        continue;
                    }

                    values[getterCall.ExpressionText] = displayValue;
                }
                catch (COMException)
                {
                }
            }

            return values;
        }

        private static Dictionary<string, string> ResolveExpressionValues(
            Debugger debugger,
            List<GetterCallToken> expressionTokens,
            InlineValueNumericDisplayMode numericDisplayMode,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds,
            IReadOnlyList<InlineValueCustomRule> customRules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            if (ShouldAbortEvaluation())
            {
                return values;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (GetterCallToken expressionToken in expressionTokens)
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                if (!seen.Add(expressionToken.ExpressionText))
                {
                    continue;
                }

                if (!HasEvaluationKind(evaluationKinds, expressionToken.EvaluationKind))
                {
                    continue;
                }

                try
                {
                    EnvDTE.Expression expression = TimedGetExpression(debugger, expressionToken.EvaluationExpressionText, 50, "Member.GetExpression");
                    if (expression == null || !expression.IsValidValue)
                    {
                        continue;
                    }

                    string normalized = NormalizeValue(expression.Value);
                    if (string.IsNullOrEmpty(normalized))
                    {
                        continue;
                    }

                    if (!TryFormatEligibleValue(
                        debugger,
                        expressionToken.EvaluationExpressionText,
                        expression.Type,
                        expression.Value,
                        normalized,
                        numericDisplayMode,
                        EnumValueRenderMode.IntegerOnly,
                        evaluationKinds,
                        ruleKinds,
                        typeRuleKinds,
                        customRules,
                        out string displayValue))
                    {
                        continue;
                    }

                    values[expressionToken.ExpressionText] = displayValue;
                }
                catch (COMException)
                {
                }
            }

            return values;
        }

        private static void AddExpressionValues(
            Dictionary<string, string> values,
            Expressions expressions,
            HashSet<string> requiredIdentifiers,
            int depth,
            Debugger debugger,
            InlineValueNumericDisplayMode numericDisplayMode,
            EnumValueRenderMode enumValueRenderMode,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds,
            IReadOnlyList<InlineValueCustomRule> customRules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (expressions == null || depth > 4 || ShouldAbortEvaluation())
            {
                return;
            }

            foreach (EnvDTE.Expression expression in expressions)
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                if (expression == null)
                {
                    continue;
                }

                try
                {
                    Stopwatch stopwatch = Stopwatch.StartNew();
                    bool isValid = expression.IsValidValue;
                    RecordPerf("Expression.IsValidValue", expression.Name, stopwatch.ElapsedMilliseconds);
                    if (!isValid)
                    {
                        continue;
                    }

                    stopwatch.Restart();
                    TryAddExpressionValue(values, expression.Name, expression.Type, expression.Value, requiredIdentifiers, debugger, numericDisplayMode, enumValueRenderMode, evaluationKinds, ruleKinds, typeRuleKinds, customRules);
                    RecordPerf("Expression.TryAddExpressionValue", expression.Name, stopwatch.ElapsedMilliseconds);
                }
                catch (COMException)
                {
                }

                try
                {
                    if (!RuntimeAllowsDataMemberRecursion())
                    {
                        continue;
                    }

                    Stopwatch stopwatch = Stopwatch.StartNew();
                    Expressions dataMembers = expression.DataMembers;
                    RecordPerf("Expression.DataMembers", expression.Name, stopwatch.ElapsedMilliseconds);
                    AddExpressionValues(values, dataMembers, requiredIdentifiers, depth + 1, debugger, numericDisplayMode, enumValueRenderMode, evaluationKinds, ruleKinds, typeRuleKinds, customRules);
                }
                catch (COMException)
                {
                }
            }
        }

        private static void TryAddExpressionValue(
            Dictionary<string, string> values,
            string name,
            string type,
            string rawValue,
            HashSet<string> requiredIdentifiers,
            Debugger debugger,
            InlineValueNumericDisplayMode numericDisplayMode,
            EnumValueRenderMode enumValueRenderMode,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds,
            IReadOnlyList<InlineValueCustomRule> customRules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (ShouldAbortEvaluation())
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            string simpleName = ExtractTrailingIdentifier(name);
            bool needsExact = requiredIdentifiers.Contains(name);
            bool needsSimple = !string.IsNullOrEmpty(simpleName) && requiredIdentifiers.Contains(simpleName);
            if (!needsExact && !needsSimple)
            {
                return;
            }

            string normalized = NormalizeValue(rawValue);
            if (string.IsNullOrEmpty(normalized))
            {
                return;
            }

            string evalExpressionName = needsExact ? name : simpleName;
            if (!TryFormatEligibleValue(
                debugger,
                evalExpressionName,
                type,
                rawValue,
                normalized,
                numericDisplayMode,
                enumValueRenderMode,
                evaluationKinds,
                ruleKinds,
                typeRuleKinds,
                customRules,
                out string displayValue))
            {
                return;
            }

            normalized = displayValue;

            if (needsExact && !values.ContainsKey(name))
            {
                values[name] = normalized;
            }

            if (needsSimple && !values.ContainsKey(simpleName))
            {
                values[simpleName] = normalized;
            }
        }

        private static string NormalizeValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            string cleaned = value
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ")
                .Trim();

            cleaned = CompactStandaloneNumericLiteral(cleaned);

            if (cleaned.Length > 80)
            {
                cleaned = cleaned.Substring(0, 77) + "...";
            }

            return cleaned;
        }

        private static string CompactStandaloneNumericLiteral(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            string token = value.Trim();
            if (token.Length == 0 || token.IndexOf(' ') >= 0)
            {
                return value;
            }

            int index = 0;
            string sign = string.Empty;
            if (token[index] == '+' || token[index] == '-')
            {
                sign = token[index].ToString();
                index++;
                if (index >= token.Length)
                {
                    return value;
                }
            }

            bool isHex = false;
            bool isBinary = false;
            string prefix = string.Empty;
            if (index + 1 < token.Length &&
                token[index] == '0' &&
                (token[index + 1] == 'x' || token[index + 1] == 'X'))
            {
                isHex = true;
                prefix = token.Substring(index, 2);
                index += 2;
            }
            else if (index + 1 < token.Length &&
                     token[index] == '0' &&
                     (token[index + 1] == 'b' || token[index + 1] == 'B'))
            {
                isBinary = true;
                prefix = token.Substring(index, 2);
                index += 2;
            }

            int digitStart = index;
            while (index < token.Length && IsDigitForLiteral(token[index], isHex, isBinary))
            {
                index++;
            }

            if (index == digitStart)
            {
                return value;
            }

            string digits = token.Substring(digitStart, index - digitStart);
            string suffix = token.Substring(index);
            if (!IsNumericSuffix(suffix))
            {
                return value;
            }

            string compactDigits = TrimLeadingZeros(digits);
            if (compactDigits.Length == 0)
            {
                compactDigits = "0";
            }

            return sign + prefix + compactDigits + suffix;
        }

        private static bool IsDigitForLiteral(char ch, bool isHex, bool isBinary)
        {
            if (isHex)
            {
                return (ch >= '0' && ch <= '9') ||
                       (ch >= 'a' && ch <= 'f') ||
                       (ch >= 'A' && ch <= 'F');
            }

            if (isBinary)
            {
                return ch == '0' || ch == '1';
            }

            return ch >= '0' && ch <= '9';
        }

        private static bool IsNumericSuffix(string suffix)
        {
            if (string.IsNullOrEmpty(suffix))
            {
                return true;
            }

            for (int i = 0; i < suffix.Length; i++)
            {
                char ch = suffix[i];
                if (ch != 'u' && ch != 'U' && ch != 'l' && ch != 'L')
                {
                    return false;
                }
            }

            return true;
        }

        private static string TrimLeadingZeros(string digits)
        {
            if (string.IsNullOrEmpty(digits))
            {
                return string.Empty;
            }

            int i = 0;
            while (i < digits.Length && digits[i] == '0')
            {
                i++;
            }

            return i >= digits.Length ? string.Empty : digits.Substring(i);
        }

        private static bool TryFormatNumericValue(
            InlineValueNumericDisplayMode numericDisplayMode,
            string type,
            string normalizedValue,
            out string displayValue)
        {
            displayValue = normalizedValue;
            if (numericDisplayMode == InlineValueNumericDisplayMode.Decimal)
            {
                return false;
            }

            if (!IsIntegralScalarType(type))
            {
                return false;
            }

            return TryFormatNumericLiteral(numericDisplayMode, normalizedValue, out displayValue);
        }

        private static bool TryFormatEnumValue(
            Debugger debugger,
            string evalExpressionName,
            string type,
            string rawValue,
            string normalizedValue,
            EnumValueRenderMode enumValueRenderMode,
            InlineValueNumericDisplayMode numericDisplayMode,
            out string displayValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            displayValue = normalizedValue;
            if (debugger == null || string.IsNullOrWhiteSpace(evalExpressionName))
            {
                return false;
            }

            bool likelyEnumType = IsDisplayableEnumType(type, rawValue, normalizedValue);
            if (!likelyEnumType && !LooksLikeEnumSymbolValue(rawValue, normalizedValue))
            {
                return false;
            }

            if (!TryReadEnumIntegerValue(debugger, evalExpressionName, out string integerText))
            {
                return false;
            }

            string formattedInteger = integerText;
            TryFormatNumericLiteral(numericDisplayMode, integerText, out formattedInteger);
            displayValue = formattedInteger;
            return true;
        }

        private static bool TryFormatNumericLiteral(
            InlineValueNumericDisplayMode numericDisplayMode,
            string value,
            out string formatted)
        {
            formatted = value;
            if (numericDisplayMode == InlineValueNumericDisplayMode.Decimal)
            {
                return false;
            }

            if (!TryParseSignedInteger(value, out long numericValue))
            {
                return false;
            }

            bool isNegative = numericValue < 0;
            ulong magnitude = isNegative
                ? (ulong)(-(numericValue + 1)) + 1UL
                : (ulong)numericValue;

            if (numericDisplayMode == InlineValueNumericDisplayMode.Hexadecimal)
            {
                string hex = "0x" + magnitude.ToString("X", CultureInfo.InvariantCulture);
                formatted = isNegative ? "-" + hex : hex;
                return true;
            }

            string binary = "0b" + Convert.ToString((long)magnitude, 2);
            formatted = isNegative ? "-" + binary : binary;
            return true;
        }

        private static bool TryParseSignedInteger(string value, out long parsed)
        {
            parsed = 0;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string trimmed = value.Trim();
            if (trimmed.Length == 0)
            {
                return false;
            }

            int end = trimmed.Length - 1;
            while (end >= 0 && char.IsWhiteSpace(trimmed[end]))
            {
                end--;
            }

            while (end >= 0)
            {
                char suffix = trimmed[end];
                if (suffix == 'u' || suffix == 'U' || suffix == 'l' || suffix == 'L')
                {
                    end--;
                    continue;
                }

                break;
            }

            if (end < 0)
            {
                return false;
            }

            string core = trimmed.Substring(0, end + 1).Trim();
            if (core.Length == 0)
            {
                return false;
            }

            bool isNegative = false;
            if (core[0] == '+' || core[0] == '-')
            {
                isNegative = core[0] == '-';
                core = core.Substring(1).Trim();
                if (core.Length == 0)
                {
                    return false;
                }
            }

            NumberStyles styles = NumberStyles.Integer;
            if (core.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                core = core.Substring(2);
                if (core.Length == 0)
                {
                    return false;
                }

                styles = NumberStyles.AllowHexSpecifier;
            }

            if (!ulong.TryParse(core, styles, CultureInfo.InvariantCulture, out ulong unsignedValue))
            {
                return false;
            }

            if (isNegative)
            {
                if (unsignedValue > 0x8000000000000000UL)
                {
                    return false;
                }

                parsed = unsignedValue == 0x8000000000000000UL
                    ? long.MinValue
                    : -(long)unsignedValue;
                return true;
            }

            if (unsignedValue > long.MaxValue)
            {
                return false;
            }

            parsed = (long)unsignedValue;
            return true;
        }

        private static bool TryReadEnumIntegerValue(Debugger debugger, string expressionText, out string integerText)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            integerText = null;
            if (debugger == null || string.IsNullOrWhiteSpace(expressionText) || ShouldAbortEvaluation())
            {
                return false;
            }

            try
            {
                string castExpression = "(int)(" + expressionText + ")";
                EnvDTE.Expression expression = TimedGetExpression(debugger, castExpression, 50, "Enum.Cast.GetExpression");
                if (expression == null || !expression.IsValidValue)
                {
                    return false;
                }

                string normalized = NormalizeValue(expression.Value);
                if (string.IsNullOrEmpty(normalized))
                {
                    return false;
                }

                string leadingToken = ExtractLeadingToken(normalized);
                if (!TryParseSignedInteger(leadingToken, out _))
                {
                    if (!TryParseSignedInteger(normalized, out _))
                    {
                        return false;
                    }

                    leadingToken = normalized;
                }

                integerText = leadingToken;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static bool IsDisplayableEnumType(string type, string rawValue, string normalizedValue)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            if (IsSuppressiblePointerType(type) || IsArrayType(type))
            {
                return false;
            }

            if (IsNonDisplayableAggregateType(type))
            {
                return false;
            }

            if (IsExplicitEnumType(type))
            {
                return true;
            }

            // Debugger type strings for enums often omit the "enum" keyword (e.g., plain type name).
            if (IsLikelyUserDefinedType(type))
            {
                return true;
            }

            return LooksLikeEnumSymbolValue(rawValue, normalizedValue);
        }

        private static bool IsExplicitEnumType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            return ContainsTypeWord(lower, "enum");
        }

        private static bool IsNonDisplayableAggregateType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            if (ContainsTypeWord(lower, "enum"))
            {
                return false;
            }

            if (ContainsTypeWord(lower, "class") || ContainsTypeWord(lower, "struct") || ContainsTypeWord(lower, "union"))
            {
                return true;
            }

            return false;
        }

        private static bool IsDisplayableScalarType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            if (IsSuppressiblePointerType(type) || IsArrayType(type))
            {
                return false;
            }

            if (IsReferenceToNonDisplayableAggregateType(type) || IsNonDisplayableAggregateType(type))
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            if (ContainsTypeWord(lower, "bool"))
            {
                return true;
            }

            if (ContainsTypeWord(lower, "float") || ContainsTypeWord(lower, "double"))
            {
                return true;
            }

            return ContainsTypeWord(lower, "char") ||
                   ContainsTypeWord(lower, "wchar_t") ||
                   ContainsTypeWord(lower, "char8_t") ||
                   ContainsTypeWord(lower, "char16_t") ||
                   ContainsTypeWord(lower, "char32_t") ||
                   ContainsTypeWord(lower, "short") ||
                   ContainsTypeWord(lower, "int") ||
                   ContainsTypeWord(lower, "long") ||
                   ContainsTypeWord(lower, "__int8") ||
                   ContainsTypeWord(lower, "__int16") ||
                   ContainsTypeWord(lower, "__int32") ||
                   ContainsTypeWord(lower, "__int64") ||
                   IsExplicitFixedWidthIntegralType(lower) ||
                   ContainsTypeWord(lower, "size_t") ||
                   ContainsTypeWord(lower, "ptrdiff_t");
        }

        private static bool IsReferenceToNonDisplayableAggregateType(string type)
        {
            if (string.IsNullOrWhiteSpace(type) || type.IndexOf('&') < 0)
            {
                return false;
            }

            return IsNonDisplayableAggregateType(type);
        }

        private static bool IsDisplayableArrayElementType(string type, string rawValue, string normalizedValue)
        {
            if (!IsArrayType(type))
            {
                return false;
            }

            string elementType = ExtractArrayElementType(type);
            if (string.IsNullOrWhiteSpace(elementType))
            {
                return false;
            }

            if (IsSuppressiblePointerType(elementType) || IsNonDisplayableAggregateType(elementType))
            {
                return false;
            }

            if (IsDisplayableScalarType(elementType))
            {
                return true;
            }

            return IsDisplayableEnumType(elementType, rawValue, normalizedValue);
        }

        private static string ExtractArrayElementType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return string.Empty;
            }

            int bracketIndex = type.IndexOf('[');
            if (bracketIndex <= 0)
            {
                return string.Empty;
            }

            return type.Substring(0, bracketIndex).Trim();
        }

        private static bool IsSuppressiblePointerType(string type)
        {
            return IsPointerType(type) || IsFunctionPointerType(type);
        }

        private static bool IsGetterSupportedType(string type, string rawValue, string normalizedValue)
        {
            if (IsSuppressiblePointerType(type) || IsArrayType(type))
            {
                return false;
            }

            return IsDisplayableScalarType(type) || IsDisplayableEnumType(type, rawValue, normalizedValue);
        }

        private static bool IsDisplayableGetterReturnType(string type, string rawValue, string normalizedValue)
        {
            return IsGetterSupportedType(type, rawValue, normalizedValue);
        }

        private static bool IsIntegralScalarType(string type)
        {
            if (!IsDisplayableScalarType(type))
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            if (ContainsTypeWord(lower, "bool"))
            {
                return false;
            }

            if (ContainsTypeWord(lower, "float") || ContainsTypeWord(lower, "double"))
            {
                return false;
            }

            string normalized = lower
                .Replace('(', ' ')
                .Replace(')', ' ')
                .Replace(',', ' ')
                .Replace('<', ' ')
                .Replace('>', ' ')
                .Replace(':', ' ');

            string[] tokens = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string token in tokens)
            {
                if (TypeNoiseTokens.Contains(token))
                {
                    continue;
                }

                if (PrimitiveTypeTokens.Contains(token))
                {
                    return ContainsTypeWord(lower, "char") ||
                           ContainsTypeWord(lower, "short") ||
                           ContainsTypeWord(lower, "int") ||
                           ContainsTypeWord(lower, "long") ||
                           ContainsTypeWord(lower, "__int8") ||
                           ContainsTypeWord(lower, "__int16") ||
                           ContainsTypeWord(lower, "__int32") ||
                           ContainsTypeWord(lower, "__int64") ||
                           IsExplicitFixedWidthIntegralType(lower) ||
                           ContainsTypeWord(lower, "size_t") ||
                           ContainsTypeWord(lower, "ptrdiff_t");
                }
            }

            return false;
        }

        private static bool IsExplicitFixedWidthIntegralType(string typeLower)
        {
            if (string.IsNullOrWhiteSpace(typeLower))
            {
                return false;
            }

            return ContainsTypeWord(typeLower, "int8_t") ||
                   ContainsTypeWord(typeLower, "int16_t") ||
                   ContainsTypeWord(typeLower, "int32_t") ||
                   ContainsTypeWord(typeLower, "int64_t") ||
                   ContainsTypeWord(typeLower, "uint8") ||
                   ContainsTypeWord(typeLower, "uint16") ||
                   ContainsTypeWord(typeLower, "uint32") ||
                   ContainsTypeWord(typeLower, "uint64") ||
                   ContainsTypeWord(typeLower, "uint8_t") ||
                   ContainsTypeWord(typeLower, "uint16_t") ||
                   ContainsTypeWord(typeLower, "uint32_t") ||
                   ContainsTypeWord(typeLower, "uint64_t");
        }

        private static bool LooksLikeEnumSymbolValue(string rawValue, string normalizedValue)
        {
            string symbol = ExtractEnumSymbolText(normalizedValue);
            if (string.IsNullOrEmpty(symbol))
            {
                symbol = ExtractEnumSymbolText(rawValue);
            }

            return !string.IsNullOrEmpty(symbol);
        }

        private static string ExtractEnumSymbolText(string value)
        {
            string normalized = NormalizeValue(value);
            if (string.IsNullOrEmpty(normalized))
            {
                return string.Empty;
            }

            string candidate = normalized.Trim();
            int braceIndex = candidate.IndexOf('{');
            if (braceIndex > 0)
            {
                candidate = candidate.Substring(0, braceIndex);
            }

            int parenIndex = candidate.IndexOf('(');
            if (parenIndex > 0)
            {
                candidate = candidate.Substring(0, parenIndex);
            }

            candidate = candidate.Trim();
            if (candidate.Length == 0)
            {
                return string.Empty;
            }

            if (TryParseSignedInteger(candidate, out _))
            {
                return string.Empty;
            }

            if (!candidate.Any(ch => char.IsLetter(ch) || ch == '_'))
            {
                return string.Empty;
            }

            return candidate;
        }

        private static bool TryFormatEligibleValue(
            Debugger debugger,
            string evalExpressionName,
            string type,
            string rawValue,
            string normalizedValue,
            InlineValueNumericDisplayMode numericDisplayMode,
            EnumValueRenderMode enumValueRenderMode,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds,
            IReadOnlyList<InlineValueCustomRule> customRules,
            out string displayValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            displayValue = normalizedValue;
            if (string.IsNullOrEmpty(normalizedValue))
            {
                return false;
            }

            CustomRuleDecision customRuleDecision = GetCustomRuleDecision(customRules, type, evalExpressionName);
            if (customRuleDecision == CustomRuleDecision.Hide)
            {
                return false;
            }

            bool forceShow = customRuleDecision == CustomRuleDecision.Show;

            if (TryFormatNullPointerValue(type, rawValue, normalizedValue, out string nullDisplay))
            {
                if (!forceShow &&
                    (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.NullPointers) ||
                    !IsDetailedPointerTypeEnabled(type, rawValue, normalizedValue, typeRuleKinds, allowUnknownKinds: true)))
                {
                    return false;
                }

                displayValue = nullDisplay;
            }
            else if (TryFormatPointerValue(
                debugger,
                evalExpressionName,
                type,
                rawValue,
                normalizedValue,
                numericDisplayMode,
                evaluationKinds,
                ruleKinds,
                typeRuleKinds,
                forceShow,
                out string pointerDisplay))
            {
                displayValue = pointerDisplay;
            }
            else if (IsSuppressiblePointerType(type))
            {
                return false;
            }
            else if (IsArrayType(type))
            {
                InlineValueTypeRuleKinds arrayRuleKind = GetArrayTypeRuleKind(type, rawValue, normalizedValue);
                if (!forceShow &&
                    !IsArrayRuleEnabled(arrayRuleKind, evaluationKinds, typeRuleKinds))
                {
                    return false;
                }

                if (!TryFormatArrayValue(debugger, evalExpressionName, type, rawValue, normalizedValue, numericDisplayMode, out string arrayDisplay))
                {
                    return false;
                }

                displayValue = arrayDisplay;
            }
            else if (TryFormatUninitializedValue(type, rawValue, normalizedValue, out string uninitializedDisplay))
            {
                if (!forceShow &&
                    (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.UninitializedValues) ||
                    !HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.BasicScalars)))
                {
                    return false;
                }

                displayValue = uninitializedDisplay;
            }
            else if (IsDisplayableEnumType(type, rawValue, normalizedValue))
            {
                if (!forceShow &&
                    (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.Enums) ||
                    !HasTypeRuleKind(typeRuleKinds, InlineValueTypeRuleKinds.EnumValues)))
                {
                    return false;
                }

                if (!TryFormatEnumValue(debugger, evalExpressionName, type, rawValue, normalizedValue, enumValueRenderMode, numericDisplayMode, out string enumDisplay))
                {
                    return false;
                }

                displayValue = enumDisplay;
            }
            else if (TryFormatNumericValue(numericDisplayMode, type, normalizedValue, out string numericDisplay))
            {
                InlineValueTypeRuleKinds scalarRuleKind = GetScalarTypeRuleKind(type);
                if (!forceShow &&
                    (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.BasicScalars) ||
                    !HasTypeRuleKind(typeRuleKinds, scalarRuleKind)))
                {
                    return false;
                }

                displayValue = numericDisplay;
            }
            else if (IsDisplayableScalarType(type))
            {
                InlineValueTypeRuleKinds scalarRuleKind = GetScalarTypeRuleKind(type);
                if (!forceShow &&
                    (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.BasicScalars) ||
                    !HasTypeRuleKind(typeRuleKinds, scalarRuleKind)))
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

            if (string.IsNullOrEmpty(displayValue))
            {
                return false;
            }

            return !ShouldSuppressValue(type, rawValue, displayValue, evaluationKinds, ruleKinds, typeRuleKinds, forceShow);
        }

        private static bool ShouldSuppressValue(
            string type,
            string rawValue,
            string normalizedValue,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds,
            bool forceShow)
        {
            if (string.IsNullOrWhiteSpace(type) || string.IsNullOrEmpty(normalizedValue))
            {
                return true;
            }

            if (normalizedValue.StartsWith(NullValueMarker, StringComparison.Ordinal))
            {
                return !forceShow &&
                    (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.NullPointers) ||
                    !IsDetailedPointerTypeEnabled(type, rawValue, normalizedValue, typeRuleKinds, allowUnknownKinds: true));
            }

            if (normalizedValue.StartsWith(UninitializedValueMarker, StringComparison.Ordinal))
            {
                return !forceShow && (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.UninitializedValues) ||
                    !HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.BasicScalars));
            }

            if (IsSuppressiblePointerType(type))
            {
                InlineValueTypeRuleKinds pointerRuleKind = GetPointerTypeRuleKind(type, rawValue, normalizedValue);
                if (pointerRuleKind == InlineValueTypeRuleKinds.None)
                {
                    return true;
                }

                return !forceShow && !IsPointerRuleEnabled(pointerRuleKind, evaluationKinds, ruleKinds, typeRuleKinds);
            }

            if (IsArrayType(type))
            {
                InlineValueTypeRuleKinds arrayRuleKind = GetArrayTypeRuleKind(type, rawValue, normalizedValue);
                return !forceShow &&
                    !IsArrayRuleEnabled(arrayRuleKind, evaluationKinds, typeRuleKinds);
            }

            if (IsDisplayableEnumType(type, null, normalizedValue))
            {
                return !forceShow &&
                    (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.Enums) ||
                    !HasTypeRuleKind(typeRuleKinds, InlineValueTypeRuleKinds.EnumValues));
            }

            if (IsDisplayableScalarType(type))
            {
                InlineValueTypeRuleKinds scalarRuleKind = GetScalarTypeRuleKind(type);
                return !forceShow &&
                    (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.BasicScalars) ||
                    !HasTypeRuleKind(typeRuleKinds, scalarRuleKind));
            }

            return true;
        }

        private static bool HasTypeRuleKind(InlineValueTypeRuleKinds typeRuleKinds, InlineValueTypeRuleKinds requiredKind)
        {
            return requiredKind != InlineValueTypeRuleKinds.None && (typeRuleKinds & requiredKind) == requiredKind;
        }

        private static bool IsBooleanScalarType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            return ContainsTypeWord(type.ToLowerInvariant(), "bool");
        }

        private static bool IsCharacterScalarType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            return ContainsTypeWord(lower, "char") ||
                   ContainsTypeWord(lower, "wchar_t") ||
                   ContainsTypeWord(lower, "char8_t") ||
                   ContainsTypeWord(lower, "char16_t") ||
                   ContainsTypeWord(lower, "char32_t");
        }

        private static bool IsFloatingPointScalarType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            return ContainsTypeWord(lower, "float") || ContainsTypeWord(lower, "double");
        }

        private static InlineValueTypeRuleKinds GetScalarTypeRuleKind(string type)
        {
            if (IsBooleanScalarType(type))
            {
                return InlineValueTypeRuleKinds.BooleanScalars;
            }

            if (IsCharacterScalarType(type))
            {
                return InlineValueTypeRuleKinds.CharacterScalars;
            }

            if (IsFloatingPointScalarType(type))
            {
                return InlineValueTypeRuleKinds.FloatingPointScalars;
            }

            if (IsIntegralScalarType(type))
            {
                return InlineValueTypeRuleKinds.IntegralScalars;
            }

            return InlineValueTypeRuleKinds.None;
        }

        private static InlineValueTypeRuleKinds GetArrayTypeRuleKind(string type, string rawValue, string normalizedValue)
        {
            if (!IsArrayType(type))
            {
                return InlineValueTypeRuleKinds.None;
            }

            string elementType = ExtractArrayElementType(type);
            if (IsBooleanScalarType(elementType))
            {
                return InlineValueTypeRuleKinds.BooleanArrays;
            }

            if (IsSignedCharArrayElementType(elementType))
            {
                return InlineValueTypeRuleKinds.SignedCharArrays;
            }

            if (IsUnsignedCharArrayElementType(elementType))
            {
                return InlineValueTypeRuleKinds.UnsignedCharArrays;
            }

            if (IsCharacterScalarType(elementType))
            {
                return InlineValueTypeRuleKinds.CharacterArrays;
            }

            if (IsFloatingPointScalarType(elementType))
            {
                return InlineValueTypeRuleKinds.FloatingPointArrays;
            }

            if (IsUnsignedIntegralArrayElementType(elementType))
            {
                return InlineValueTypeRuleKinds.UnsignedIntegralArrays;
            }

            if (IsSignedIntegralArrayElementType(elementType))
            {
                return InlineValueTypeRuleKinds.SignedIntegralArrays;
            }

            if (IsIntegralScalarType(elementType))
            {
                return InlineValueTypeRuleKinds.IntegralArrays;
            }

            if (IsDisplayableEnumType(elementType, rawValue, normalizedValue))
            {
                return InlineValueTypeRuleKinds.EnumArrays;
            }

            return InlineValueTypeRuleKinds.None;
        }

        private static bool IsArrayRuleEnabled(
            InlineValueTypeRuleKinds arrayRuleKind,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueTypeRuleKinds typeRuleKinds)
        {
            if (!HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.Arrays))
            {
                return false;
            }

            switch (arrayRuleKind)
            {
                case InlineValueTypeRuleKinds.SignedCharArrays:
                case InlineValueTypeRuleKinds.UnsignedCharArrays:
                    return HasTypeRuleKind(typeRuleKinds, InlineValueTypeRuleKinds.CharacterArrays) &&
                           HasTypeRuleKind(typeRuleKinds, arrayRuleKind);
                case InlineValueTypeRuleKinds.SignedIntegralArrays:
                case InlineValueTypeRuleKinds.UnsignedIntegralArrays:
                    return HasTypeRuleKind(typeRuleKinds, InlineValueTypeRuleKinds.IntegralArrays) &&
                           HasTypeRuleKind(typeRuleKinds, arrayRuleKind);
                default:
                    return HasTypeRuleKind(typeRuleKinds, arrayRuleKind);
            }
        }

        private static InlineValueTypeRuleKinds GetPointerTypeRuleKind(string type, string rawValue, string normalizedValue)
        {
            if (!IsPointerType(type) || IsFunctionPointerType(type))
            {
                return InlineValueTypeRuleKinds.None;
            }

            string elementType = ExtractPointerElementType(type);
            if (IsExplicitUnsignedPointerType(elementType, 8))
            {
                return InlineValueTypeRuleKinds.Unsigned8BitPointers;
            }

            if (IsExplicitUnsignedPointerType(elementType, 16))
            {
                return InlineValueTypeRuleKinds.Unsigned16BitPointers;
            }

            if (IsExplicitUnsignedPointerType(elementType, 32))
            {
                return InlineValueTypeRuleKinds.Unsigned32BitPointers;
            }

            if (IsCharPointerOrArrayType(type) || LooksLikeStringPointerOrArray(type, rawValue, normalizedValue))
            {
                return InlineValueTypeRuleKinds.CharacterPointers;
            }

            if (IsBooleanScalarType(elementType))
            {
                return InlineValueTypeRuleKinds.BooleanPointers;
            }

            if (IsFloatingPointScalarType(elementType))
            {
                return InlineValueTypeRuleKinds.FloatingPointPointers;
            }

            if (IsDisplayableIntegralPointerType(type))
            {
                return InlineValueTypeRuleKinds.IntegralPointers;
            }

            if (IsExplicitEnumType(elementType))
            {
                return InlineValueTypeRuleKinds.EnumPointers;
            }

            if (IsLikelyClassOrStructPointerType(type))
            {
                return InlineValueTypeRuleKinds.StructPointers;
            }

            return InlineValueTypeRuleKinds.None;
        }

        private static bool IsPointerRuleEnabled(
            InlineValueTypeRuleKinds pointerRuleKind,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds)
        {
            if (!HasTypeRuleKind(typeRuleKinds, pointerRuleKind))
            {
                return false;
            }

            switch (pointerRuleKind)
            {
                case InlineValueTypeRuleKinds.IntegralPointers:
                case InlineValueTypeRuleKinds.Unsigned8BitPointers:
                case InlineValueTypeRuleKinds.Unsigned16BitPointers:
                case InlineValueTypeRuleKinds.Unsigned32BitPointers:
                    return HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.BasicScalars) &&
                           HasRuleKind(ruleKinds, InlineValueRuleKinds.IntegralPointers);
                case InlineValueTypeRuleKinds.BooleanPointers:
                case InlineValueTypeRuleKinds.FloatingPointPointers:
                    return HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.BasicScalars);
                case InlineValueTypeRuleKinds.EnumPointers:
                    return HasEvaluationKind(evaluationKinds, InlineValueEvaluationKinds.Enums);
                case InlineValueTypeRuleKinds.CharacterPointers:
                case InlineValueTypeRuleKinds.StructPointers:
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsDetailedPointerTypeEnabled(
            string type,
            string rawValue,
            string normalizedValue,
            InlineValueTypeRuleKinds typeRuleKinds,
            bool allowUnknownKinds)
        {
            InlineValueTypeRuleKinds pointerRuleKind = GetPointerTypeRuleKind(type, rawValue, normalizedValue);
            return pointerRuleKind == InlineValueTypeRuleKinds.None
                ? allowUnknownKinds
                : HasTypeRuleKind(typeRuleKinds, pointerRuleKind);
        }

        private static bool IsSignedCharArrayElementType(string elementType)
        {
            if (string.IsNullOrWhiteSpace(elementType))
            {
                return false;
            }

            string lower = elementType.ToLowerInvariant();
            return (ContainsTypeWord(lower, "signed") && ContainsTypeWord(lower, "char")) ||
                   ContainsTypeWord(lower, "int8_t");
        }

        private static bool IsUnsignedCharArrayElementType(string elementType)
        {
            if (string.IsNullOrWhiteSpace(elementType))
            {
                return false;
            }

            string lower = elementType.ToLowerInvariant();
            return (ContainsTypeWord(lower, "unsigned") && ContainsTypeWord(lower, "char")) ||
                   ContainsTypeWord(lower, "uint8") ||
                   ContainsTypeWord(lower, "uint8_t") ||
                   ContainsTypeWord(lower, "byte");
        }

        private static bool IsUnsignedIntegralArrayElementType(string elementType)
        {
            if (!IsIntegralScalarType(elementType) || string.IsNullOrWhiteSpace(elementType))
            {
                return false;
            }

            if (IsSignedCharArrayElementType(elementType) || IsUnsignedCharArrayElementType(elementType))
            {
                return false;
            }

            string lower = elementType.ToLowerInvariant();
            return ContainsTypeWord(lower, "unsigned") ||
                   ContainsTypeWord(lower, "uint") ||
                   ContainsTypeWord(lower, "size_t") ||
                   ContainsTypeWord(lower, "byte") ||
                   ContainsTypeWord(lower, "word") ||
                   ContainsTypeWord(lower, "dword");
        }

        private static bool IsSignedIntegralArrayElementType(string elementType)
        {
            if (!IsIntegralScalarType(elementType) || string.IsNullOrWhiteSpace(elementType))
            {
                return false;
            }

            if (IsSignedCharArrayElementType(elementType) || IsUnsignedCharArrayElementType(elementType))
            {
                return false;
            }

            return !IsUnsignedIntegralArrayElementType(elementType);
        }

        private static bool IsArrayType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            return type.IndexOf('[') >= 0 && type.IndexOf(']') >= 0;
        }

        private static bool IsPointerType(string type)
        {
            return !string.IsNullOrWhiteSpace(type) && type.IndexOf('*') >= 0;
        }

        private static string ExtractPointerElementType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return string.Empty;
            }

            int pointerIndex = type.IndexOf('*');
            if (pointerIndex <= 0)
            {
                return string.Empty;
            }

            return type.Substring(0, pointerIndex).Trim();
        }

        private static bool IsExplicitUnsignedPointerType(string elementType, int bitWidth)
        {
            if (string.IsNullOrWhiteSpace(elementType))
            {
                return false;
            }

            string lower = elementType.ToLowerInvariant();
            switch (bitWidth)
            {
                case 8:
                    return ContainsTypeWord(lower, "uint8") ||
                           ContainsTypeWord(lower, "uint8_t") ||
                           ContainsTypeWord(lower, "byte") ||
                           (ContainsTypeWord(lower, "unsigned") &&
                            (ContainsTypeWord(lower, "char") || ContainsTypeWord(lower, "__int8")));
                case 16:
                    return ContainsTypeWord(lower, "uint16") ||
                           ContainsTypeWord(lower, "uint16_t") ||
                           ContainsTypeWord(lower, "word") ||
                           (ContainsTypeWord(lower, "unsigned") &&
                            (ContainsTypeWord(lower, "short") || ContainsTypeWord(lower, "__int16")));
                case 32:
                    return ContainsTypeWord(lower, "uint32") ||
                           ContainsTypeWord(lower, "uint32_t") ||
                           ContainsTypeWord(lower, "dword") ||
                           (ContainsTypeWord(lower, "unsigned") &&
                            (ContainsTypeWord(lower, "int") ||
                             ContainsTypeWord(lower, "long") ||
                             ContainsTypeWord(lower, "__int32")));
                default:
                    return false;
            }
        }

        private static bool IsDisplayableIntegralPointerType(string type)
        {
            if (!IsPointerType(type) || IsFunctionPointerType(type))
            {
                return false;
            }

            string elementType = ExtractPointerElementType(type);
            if (string.IsNullOrWhiteSpace(elementType))
            {
                return false;
            }

            string lower = elementType.ToLowerInvariant();
            if (ContainsTypeWord(lower, "char") ||
                ContainsTypeWord(lower, "wchar_t") ||
                ContainsTypeWord(lower, "char8_t") ||
                ContainsTypeWord(lower, "char16_t") ||
                ContainsTypeWord(lower, "char32_t") ||
                ContainsTypeWord(lower, "bool"))
            {
                return false;
            }

            return IsIntegralScalarType(elementType) ||
                   ContainsTypeWord(lower, "uint") ||
                   ContainsTypeWord(lower, "ushort") ||
                   ContainsTypeWord(lower, "ulong") ||
                   ContainsTypeWord(lower, "byte") ||
                   ContainsTypeWord(lower, "word") ||
                   ContainsTypeWord(lower, "dword");
        }

        private static bool TryFormatNullPointerValue(string type, string rawValue, string normalizedValue, out string displayValue)
        {
            displayValue = normalizedValue;
            if (!IsPointerType(type) || IsFunctionPointerType(type))
            {
                return false;
            }

            if (!IsNullLikeValue(rawValue) && !IsNullLikeValue(normalizedValue))
            {
                return false;
            }

            displayValue = NullValueMarker + "null";
            return true;
        }

        private static bool TryFormatPointerValue(
            Debugger debugger,
            string expressionName,
            string type,
            string rawValue,
            string normalizedValue,
            InlineValueNumericDisplayMode numericDisplayMode,
            InlineValueEvaluationKinds evaluationKinds,
            InlineValueRuleKinds ruleKinds,
            InlineValueTypeRuleKinds typeRuleKinds,
            bool forceShow,
            out string displayValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            displayValue = normalizedValue;
            InlineValueTypeRuleKinds pointerRuleKind = GetPointerTypeRuleKind(type, rawValue, normalizedValue);
            if (pointerRuleKind == InlineValueTypeRuleKinds.None ||
                string.IsNullOrWhiteSpace(expressionName) ||
                debugger == null)
            {
                return false;
            }

            if (IsNullLikeValue(rawValue) || IsNullLikeValue(normalizedValue))
            {
                return false;
            }

            if (!forceShow && !IsPointerRuleEnabled(pointerRuleKind, evaluationKinds, ruleKinds, typeRuleKinds))
            {
                return false;
            }

            string pointerDisplay = BuildPointerDisplay(rawValue, normalizedValue);
            if (pointerRuleKind == InlineValueTypeRuleKinds.CharacterPointers)
            {
                if (!TryFormatCharPointerOrArrayValue(debugger, expressionName, type, rawValue, normalizedValue, out string stringPointerDisplay))
                {
                    return false;
                }

                displayValue = pointerDisplay + " (" + StripDisplayMarker(stringPointerDisplay) + ")";
                return true;
            }

            string dereferenceExpression = "*(" + expressionName + ")";
            EnvDTE.Expression dereferencedExpression;
            try
            {
                dereferencedExpression = TimedGetExpression(debugger, dereferenceExpression, 50, "Pointer.Deref.GetExpression");
            }
            catch (COMException)
            {
                return false;
            }

            if (dereferencedExpression == null || !dereferencedExpression.IsValidValue)
            {
                return false;
            }

            string normalizedDerefValue = NormalizeValue(dereferencedExpression.Value);
            if (string.IsNullOrEmpty(normalizedDerefValue))
            {
                return false;
            }

            if (pointerRuleKind == InlineValueTypeRuleKinds.StructPointers)
            {
                if (!TryFormatAggregatePointerValue(dereferencedExpression.Value, normalizedDerefValue, out string aggregateDisplay))
                {
                    return false;
                }

                displayValue = pointerDisplay + " (" + aggregateDisplay + ")";
                return true;
            }

            if (!TryFormatEligibleValue(
                debugger,
                dereferenceExpression,
                dereferencedExpression.Type,
                dereferencedExpression.Value,
                normalizedDerefValue,
                numericDisplayMode,
                EnumValueRenderMode.IntegerOnly,
                evaluationKinds,
                ruleKinds,
                typeRuleKinds,
                null,
                out string dereferencedDisplay))
            {
                return false;
            }

            displayValue = pointerDisplay + " (" + StripDisplayMarker(dereferencedDisplay) + ")";
            return true;
        }

        private static string BuildPointerDisplay(string rawValue, string normalizedValue)
        {
            string pointerDisplay = ExtractLeadingToken(normalizedValue);
            if (string.IsNullOrWhiteSpace(pointerDisplay) || !IsAddressLikeValue(pointerDisplay))
            {
                pointerDisplay = ExtractLeadingToken(rawValue);
            }

            if (string.IsNullOrWhiteSpace(pointerDisplay))
            {
                pointerDisplay = normalizedValue.Trim();
            }

            return pointerDisplay;
        }

        private static bool TryFormatAggregatePointerValue(string rawValue, string normalizedValue, out string displayValue)
        {
            displayValue = normalizedValue;
            string summary = NormalizeValue(normalizedValue) ?? NormalizeValue(rawValue);
            if (string.IsNullOrWhiteSpace(summary))
            {
                return false;
            }

            if (summary.Length > 48)
            {
                summary = summary.Substring(0, 45) + "...";
            }

            displayValue = summary;
            return true;
        }

        private static bool TryFormatUninitializedValue(string type, string rawValue, string normalizedValue, out string displayValue)
        {
            displayValue = normalizedValue;
            if (string.IsNullOrWhiteSpace(normalizedValue) && string.IsNullOrWhiteSpace(rawValue))
            {
                return false;
            }

            if (ContainsUninitializedText(rawValue) || ContainsUninitializedText(normalizedValue))
            {
                displayValue = UninitializedValueMarker + "Not Init";
                return true;
            }

            if (!IsDisplayableScalarType(type))
            {
                return false;
            }

            if (!IsLikelyUninitializedPattern(normalizedValue))
            {
                return false;
            }

            displayValue = UninitializedValueMarker + "Not Init";
            return true;
        }

        private static bool ContainsUninitializedText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string lower = value.ToLowerInvariant();
            return lower.Contains("uninitialized") ||
                   lower.Contains("not initialized") ||
                   lower.Contains("notinitialized") ||
                   lower.Contains("uninit") ||
                   string.Equals(lower.Trim(), "???", StringComparison.Ordinal);
        }

        private static bool IsNullLikeValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string trimmed = value.Trim();
            if (string.Equals(trimmed, "nullptr", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "null", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "(nil)", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string token = ExtractLeadingToken(trimmed);
            if (string.Equals(token, "0", StringComparison.Ordinal))
            {
                return true;
            }

            if (token.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                string hex = StripHexDecorators(token.Substring(2));
                return hex.Length > 0 && IsAllSameHexDigit(hex, '0');
            }

            string plain = StripHexDecorators(token);
            return plain.Length >= 4 && IsAllSameHexDigit(plain, '0');
        }

        private static bool IsLikelyUninitializedPattern(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string trimmed = value.Trim();
            if (string.Equals(trimmed, "-858993460", StringComparison.Ordinal) ||
                string.Equals(trimmed, "3435973836", StringComparison.Ordinal) ||
                string.Equals(trimmed, "-842150451", StringComparison.Ordinal) ||
                string.Equals(trimmed, "3452816845", StringComparison.Ordinal))
            {
                return true;
            }

            string token = ExtractLeadingToken(trimmed);
            bool isNegative = false;
            if (token.StartsWith("-", StringComparison.Ordinal))
            {
                isNegative = true;
                token = token.Substring(1);
            }
            else if (token.StartsWith("+", StringComparison.Ordinal))
            {
                token = token.Substring(1);
            }

            if (token.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                token = token.Substring(2);
            }

            token = StripHexDecorators(token).ToLowerInvariant();
            if (token.Length >= 8 && (IsAllSameHexDigit(token, 'c') || IsRepeatedHexPair(token, 'c', 'd')))
            {
                return true;
            }

            if (TryParseSignedInteger(trimmed, out long parsed))
            {
                if (parsed == -858993460L || parsed == -842150451L || parsed == -3689348814741910324L)
                {
                    return true;
                }
            }

            if (!isNegative && ulong.TryParse(token, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out ulong unsignedHex))
            {
                return unsignedHex == 0xCCCCCCCCUL ||
                       unsignedHex == 0xCDCDCDCDUL ||
                       unsignedHex == 0xCCCCCCCCCCCCCCCCUL ||
                       unsignedHex == 0xCDCDCDCDCDCDCDCDUL;
            }

            return false;
        }

        private static string ExtractLeadingToken(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            int i = 0;
            while (i < value.Length && !char.IsWhiteSpace(value[i]) && value[i] != ',')
            {
                i++;
            }

            return value.Substring(0, i);
        }

        private static string StripHexDecorators(string token)
        {
            if (string.IsNullOrEmpty(token))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(token.Length);
            for (int i = 0; i < token.Length; i++)
            {
                char ch = token[i];
                if (ch == '`' || ch == '\'' || ch == '_')
                {
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }

        private static bool IsAllSameHexDigit(string token, char digit)
        {
            if (string.IsNullOrEmpty(token))
            {
                return false;
            }

            for (int i = 0; i < token.Length; i++)
            {
                if (char.ToLowerInvariant(token[i]) != char.ToLowerInvariant(digit))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsRepeatedHexPair(string token, char first, char second)
        {
            if (string.IsNullOrEmpty(token) || (token.Length % 2) != 0)
            {
                return false;
            }

            char lowerFirst = char.ToLowerInvariant(first);
            char lowerSecond = char.ToLowerInvariant(second);
            for (int i = 0; i < token.Length; i += 2)
            {
                if (char.ToLowerInvariant(token[i]) != lowerFirst || char.ToLowerInvariant(token[i + 1]) != lowerSecond)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryFormatArrayValue(
            Debugger debugger,
            string expressionName,
            string type,
            string rawValue,
            string normalizedValue,
            InlineValueNumericDisplayMode numericDisplayMode,
            out string displayValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            displayValue = normalizedValue;
            if (!IsArrayType(type))
            {
                return false;
            }

            if (!IsDisplayableArrayElementType(type, rawValue, normalizedValue))
            {
                return false;
            }

            if (IsCharacterArrayType(type) &&
                (TryExtractQuotedStringContent(rawValue, out string content) || TryExtractQuotedStringContent(normalizedValue, out content)))
            {
                List<string> entriesFromString = BuildCharArrayEntries(content, 3, out bool hasMoreFromString);
                if (entriesFromString.Count > 0)
                {
                    displayValue = FormatArrayEntries(entriesFromString, hasMoreFromString, numericDisplayMode);
                    return true;
                }
            }

            if (TryExtractArrayEntries(rawValue, out List<string> fromRaw, out bool fromRawHasMore))
            {
                displayValue = FormatArrayEntries(fromRaw, fromRawHasMore, numericDisplayMode);
                return true;
            }

            if (TryExtractArrayEntries(normalizedValue, out List<string> fromNormalized, out bool fromNormalizedHasMore))
            {
                displayValue = FormatArrayEntries(fromNormalized, fromNormalizedHasMore, numericDisplayMode);
                return true;
            }

            if (RuntimeAllowsArrayProbing() &&
                TryReadArrayEntriesFromDebugger(debugger, expressionName, 3, out List<string> fromDebugger, out bool fromDebuggerHasMore))
            {
                displayValue = FormatArrayEntries(fromDebugger, fromDebuggerHasMore, numericDisplayMode);
                return true;
            }

            return false;
        }

        private static bool IsCharacterArrayType(string type)
        {
            string elementType = ExtractArrayElementType(type).ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(elementType))
            {
                return false;
            }

            return ContainsTypeWord(elementType, "char") ||
                   ContainsTypeWord(elementType, "wchar_t") ||
                   ContainsTypeWord(elementType, "char8_t") ||
                   ContainsTypeWord(elementType, "char16_t") ||
                   ContainsTypeWord(elementType, "char32_t");
        }

        private static List<string> BuildCharArrayEntries(string content, int maxEntries, out bool hasMore)
        {
            var entries = new List<string>(maxEntries);
            hasMore = false;
            if (string.IsNullOrEmpty(content) || maxEntries <= 0)
            {
                return entries;
            }

            int limit = Math.Min(content.Length, maxEntries);
            for (int i = 0; i < limit; i++)
            {
                entries.Add("'" + EscapeCharForChip(content[i]) + "'");
            }

            hasMore = content.Length > maxEntries;
            return entries;
        }

        private static string EscapeCharForChip(char ch)
        {
            switch (ch)
            {
                case '\0':
                    return "\\0";
                case '\n':
                    return "\\n";
                case '\r':
                    return "\\r";
                case '\t':
                    return "\\t";
                case '\'':
                    return "\\'";
                case '\\':
                    return "\\\\";
                default:
                    return ch.ToString();
            }
        }

        private static string FormatArrayEntries(
            List<string> entries,
            bool hasMore,
            InlineValueNumericDisplayMode numericDisplayMode)
        {
            if (entries == null || entries.Count == 0)
            {
                return "...";
            }

            var displayEntries = new List<string>(entries.Count);
            foreach (string entry in entries)
            {
                if (TryFormatNumericLiteral(numericDisplayMode, entry, out string converted))
                {
                    displayEntries.Add(converted);
                }
                else
                {
                    displayEntries.Add(entry);
                }
            }

            string joined = string.Join(", ", displayEntries);
            return hasMore ? joined + ", ..." : joined;
        }

        private static bool TryExtractArrayEntries(string value, out List<string> entries, out bool hasMore)
        {
            entries = new List<string>(3);
            hasMore = false;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            int openBrace = value.IndexOf('{');
            if (openBrace < 0)
            {
                return false;
            }

            int closeBrace = FindMatchingBrace(value, openBrace);
            if (closeBrace <= openBrace)
            {
                return false;
            }

            string body = value.Substring(openBrace + 1, closeBrace - openBrace - 1);
            ParseArrayEntryTokens(body, entries, out hasMore);
            return entries.Count > 0;
        }

        private static int FindMatchingBrace(string value, int openBrace)
        {
            int depth = 0;
            bool inQuoted = false;
            char quote = '\0';
            bool escaped = false;
            for (int i = openBrace; i < value.Length; i++)
            {
                char ch = value[i];
                if (inQuoted)
                {
                    if (escaped)
                    {
                        escaped = false;
                        continue;
                    }

                    if (ch == '\\')
                    {
                        escaped = true;
                        continue;
                    }

                    if (ch == quote)
                    {
                        inQuoted = false;
                    }

                    continue;
                }

                if (ch == '"' || ch == '\'')
                {
                    inQuoted = true;
                    quote = ch;
                    escaped = false;
                    continue;
                }

                if (ch == '{')
                {
                    depth++;
                    continue;
                }

                if (ch == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        private static void ParseArrayEntryTokens(string body, List<string> entries, out bool hasMore)
        {
            hasMore = false;
            if (string.IsNullOrEmpty(body))
            {
                return;
            }

            var tokenBuilder = new StringBuilder(body.Length);
            int nestedDepth = 0;
            bool inQuoted = false;
            char quote = '\0';
            bool escaped = false;
            for (int i = 0; i < body.Length; i++)
            {
                char ch = body[i];
                if (inQuoted)
                {
                    tokenBuilder.Append(ch);
                    if (escaped)
                    {
                        escaped = false;
                        continue;
                    }

                    if (ch == '\\')
                    {
                        escaped = true;
                        continue;
                    }

                    if (ch == quote)
                    {
                        inQuoted = false;
                    }

                    continue;
                }

                if (ch == '"' || ch == '\'')
                {
                    inQuoted = true;
                    quote = ch;
                    escaped = false;
                    tokenBuilder.Append(ch);
                    continue;
                }

                if (ch == '{' || ch == '(' || ch == '[' || ch == '<')
                {
                    nestedDepth++;
                    tokenBuilder.Append(ch);
                    continue;
                }

                if (ch == '}' || ch == ')' || ch == ']' || ch == '>')
                {
                    if (nestedDepth > 0)
                    {
                        nestedDepth--;
                    }

                    tokenBuilder.Append(ch);
                    continue;
                }

                if (ch == ',' && nestedDepth == 0)
                {
                    AddArrayEntryToken(tokenBuilder, entries, ref hasMore);
                    continue;
                }

                tokenBuilder.Append(ch);
            }

            AddArrayEntryToken(tokenBuilder, entries, ref hasMore);
        }

        private static void AddArrayEntryToken(StringBuilder tokenBuilder, List<string> entries, ref bool hasMore)
        {
            string token = tokenBuilder.ToString();
            tokenBuilder.Clear();
            string normalizedEntry = NormalizeArrayEntry(token);
            if (string.IsNullOrEmpty(normalizedEntry))
            {
                return;
            }

            if (entries.Count < 3)
            {
                entries.Add(normalizedEntry);
                return;
            }

            hasMore = true;
        }

        private static string NormalizeArrayEntry(string value)
        {
            string normalized = NormalizeValue(value);
            if (string.IsNullOrEmpty(normalized))
            {
                return string.Empty;
            }

            if (IsAddressLikeValue(normalized))
            {
                return string.Empty;
            }

            if (normalized.Length > 24)
            {
                normalized = normalized.Substring(0, 21) + "...";
            }

            return normalized;
        }

        private static bool TryReadArrayEntriesFromDebugger(
            Debugger debugger,
            string expressionName,
            int maxEntries,
            out List<string> entries,
            out bool hasMore)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            entries = new List<string>(maxEntries);
            hasMore = false;
            if (debugger == null || string.IsNullOrWhiteSpace(expressionName) || maxEntries <= 0 || ShouldAbortEvaluation())
            {
                return false;
            }

            foreach (string candidate in BuildIndexExpressionCandidates(expressionName))
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                if (TryReadArrayEntriesFromCandidate(debugger, candidate, maxEntries, out entries, out hasMore))
                {
                    return true;
                }
            }

            entries = new List<string>(maxEntries);
            hasMore = false;
            return false;
        }

        private static bool TryReadArrayEntriesFromCandidate(
            Debugger debugger,
            string candidateExpression,
            int maxEntries,
            out List<string> entries,
            out bool hasMore)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            entries = new List<string>(maxEntries);
            hasMore = false;
            if (string.IsNullOrWhiteSpace(candidateExpression))
            {
                return false;
            }

            bool sawAnyIndex = false;
            int probeLimit = maxEntries + 8;
            for (int i = 0; i < probeLimit; i++)
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                string indexedExpression = candidateExpression + "[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                EnvDTE.Expression valueExpr;
                try
                {
                    valueExpr = TimedGetExpression(debugger, indexedExpression, 50, "Array.Index.GetExpression");
                }
                catch (COMException)
                {
                    break;
                }

                if (valueExpr == null || !valueExpr.IsValidValue)
                {
                    break;
                }

                sawAnyIndex = true;
                string normalizedEntry = NormalizeArrayEntry(valueExpr.Value);
                if (string.IsNullOrEmpty(normalizedEntry))
                {
                    continue;
                }

                if (entries.Count < maxEntries)
                {
                    entries.Add(normalizedEntry);
                    continue;
                }

                hasMore = true;
                break;
            }

            return sawAnyIndex;
        }

        private static bool IsCharPointerOrArrayType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            bool hasPointerOrArray = type.IndexOf('*') >= 0 || (type.IndexOf('[') >= 0 && type.IndexOf(']') >= 0);
            if (!hasPointerOrArray)
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            return ContainsTypeWord(lower, "char") ||
                   ContainsTypeWord(lower, "wchar_t") ||
                   ContainsTypeWord(lower, "char8_t") ||
                   ContainsTypeWord(lower, "char16_t") ||
                   ContainsTypeWord(lower, "char32_t");
        }

        private static bool TryFormatCharPointerOrArrayValue(string type, string rawValue, string normalizedValue, out string displayValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return TryFormatCharPointerOrArrayValue(null, null, type, rawValue, normalizedValue, out displayValue);
        }

        private static bool TryFormatCharPointerOrArrayValue(
            Debugger debugger,
            string expressionName,
            string type,
            string rawValue,
            string normalizedValue,
            out string displayValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            displayValue = normalizedValue;
            if (!IsCharPointerOrArrayType(type) && !LooksLikeStringPointerOrArray(type, rawValue, normalizedValue))
            {
                return false;
            }

            if (!TryExtractQuotedStringContent(rawValue, out string content) &&
                !TryExtractQuotedStringContent(normalizedValue, out content) &&
                !TryExtractCharsFromSingleQuotedItems(rawValue, out content) &&
                !TryExtractCharsFromSingleQuotedItems(normalizedValue, out content) &&
                (!RuntimeAllowsCharProbing() || !ThreadCheckAndTryReadChars(debugger, expressionName, out content)))
            {
                displayValue = "...";
                return true;
            }

            if (content.Length >= 4)
            {
                displayValue = "\"" + content.Substring(0, 4) + "...\"";
                return true;
            }

            displayValue = "\"" + content + "...\"";
            return true;
        }

        private static bool LooksLikeStringPointerOrArray(string type, string rawValue, string normalizedValue)
        {
            bool hasPointerOrArray = !string.IsNullOrWhiteSpace(type) &&
                                     (type.IndexOf('*') >= 0 || (type.IndexOf('[') >= 0 && type.IndexOf(']') >= 0));

            if (TryExtractQuotedStringContent(rawValue, out _) || TryExtractQuotedStringContent(normalizedValue, out _))
            {
                return true;
            }

            if (!hasPointerOrArray)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(normalizedValue))
            {
                return false;
            }

            string trimmed = normalizedValue.Trim();
            return string.Equals(trimmed, "...", StringComparison.Ordinal);
        }

        private static bool ThreadCheckAndTryReadChars(Debugger debugger, string expressionName, out string content)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return TryReadCharsFromDebugger(debugger, expressionName, out content);
        }

        private static bool TryReadCharsFromDebugger(Debugger debugger, string expressionName, out string content)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            content = string.Empty;
            if (debugger == null || string.IsNullOrWhiteSpace(expressionName) || ShouldAbortEvaluation())
            {
                return false;
            }

            foreach (string candidate in BuildIndexExpressionCandidates(expressionName))
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                if (TryReadCharsFromCandidate(debugger, candidate, out string fromCandidate))
                {
                    content = fromCandidate;
                    return true;
                }
            }

            return false;
        }

        private static bool TryReadCharsFromCandidate(Debugger debugger, string candidateExpression, out string content)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            content = string.Empty;
            if (string.IsNullOrWhiteSpace(candidateExpression))
            {
                return false;
            }

            if (TryReadCharsWithStringFormat(debugger, candidateExpression, out string fromStringFormat))
            {
                content = fromStringFormat;
                return true;
            }

            var chars = new List<char>(4);
            for (int i = 0; i < 4; i++)
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                string elementExpr = candidateExpression + "[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                EnvDTE.Expression valueExpr;
                try
                {
                    valueExpr = TimedGetExpression(debugger, elementExpr, 50, "CharArray.Index.GetExpression");
                }
                catch (COMException)
                {
                    break;
                }

                if (valueExpr == null || !valueExpr.IsValidValue)
                {
                    break;
                }

                if (!TryParseCharValue(valueExpr.Value, out char ch, out bool isTerminator))
                {
                    break;
                }

                if (isTerminator)
                {
                    break;
                }

                chars.Add(ch);
            }

            if (chars.Count == 0)
            {
                return false;
            }

            content = new string(chars.ToArray());
            return true;
        }

        private static bool TryReadCharsWithStringFormat(Debugger debugger, string candidateExpression, out string content)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            content = string.Empty;
            if (debugger == null || string.IsNullOrWhiteSpace(candidateExpression) || ShouldAbortEvaluation())
            {
                return false;
            }

            string[] probes =
            {
                candidateExpression + ",s",
                candidateExpression + ",sb",
                candidateExpression + ",su"
            };

            foreach (string probe in probes)
            {
                if (ShouldAbortEvaluation())
                {
                    break;
                }

                EnvDTE.Expression expression;
                try
                {
                    expression = TimedGetExpression(debugger, probe, 50, "CharArray.Probe.GetExpression");
                }
                catch (COMException)
                {
                    continue;
                }

                if (expression == null || !expression.IsValidValue)
                {
                    continue;
                }

                if (TryExtractQuotedStringContent(expression.Value, out string fromRaw) && !string.IsNullOrEmpty(fromRaw))
                {
                    content = fromRaw;
                    return true;
                }

                string normalized = NormalizeValue(expression.Value);
                if (TryExtractQuotedStringContent(normalized, out string fromNormalized) && !string.IsNullOrEmpty(fromNormalized))
                {
                    content = fromNormalized;
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<string> BuildIndexExpressionCandidates(string expressionName)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            string raw = expressionName.Trim();
            if (seen.Add(raw))
            {
                yield return raw;
            }

            int bracketIndex = raw.IndexOf('[');
            if (bracketIndex > 0)
            {
                string baseExpr = raw.Substring(0, bracketIndex).Trim();
                if (!string.IsNullOrEmpty(baseExpr) && seen.Add(baseExpr))
                {
                    yield return baseExpr;
                }
            }

            int spaceIndex = raw.IndexOf(' ');
            if (spaceIndex > 0)
            {
                string noSuffix = raw.Substring(0, spaceIndex).Trim();
                if (!string.IsNullOrEmpty(noSuffix) && seen.Add(noSuffix))
                {
                    yield return noSuffix;
                }
            }

            string trailing = ExtractTrailingIdentifier(raw);
            if (!string.IsNullOrEmpty(trailing) && seen.Add(trailing))
            {
                yield return trailing;
            }
        }

        private static bool TryParseCharValue(string value, out char ch, out bool isTerminator)
        {
            ch = '\0';
            isTerminator = false;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (TryExtractCharsFromSingleQuotedItems(value, out string quoted) && quoted.Length > 0)
            {
                ch = quoted[0];
                isTerminator = ch == '\0';
                return true;
            }

            string trimmed = value.Trim();
            int index = 0;
            while (index < trimmed.Length && char.IsWhiteSpace(trimmed[index]))
            {
                index++;
            }

            if (index >= trimmed.Length)
            {
                return false;
            }

            int start = index;
            if (trimmed[index] == '+' || trimmed[index] == '-')
            {
                index++;
            }

            bool hex = index + 1 < trimmed.Length && trimmed[index] == '0' && (trimmed[index + 1] == 'x' || trimmed[index + 1] == 'X');
            if (hex)
            {
                index += 2;
                int hexStart = index;
                while (index < trimmed.Length && Uri.IsHexDigit(trimmed[index]))
                {
                    index++;
                }

                if (index == hexStart)
                {
                    return false;
                }

                string token = trimmed.Substring(start, index - start);
                if (!long.TryParse(token, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out long hexValue))
                {
                    // signed hex with prefix isn't parsed by AllowHexSpecifier when token has 0x.
                    string stripped = token.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? token.Substring(2) : token;
                    if (!long.TryParse(stripped, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out hexValue))
                    {
                        return false;
                    }
                }

                if (hexValue < 0 || hexValue > 255)
                {
                    return false;
                }

                ch = (char)hexValue;
                isTerminator = hexValue == 0;
                return true;
            }
            else
            {
                while (index < trimmed.Length && char.IsDigit(trimmed[index]))
                {
                    index++;
                }

                if (index == start || (index == start + 1 && (trimmed[start] == '+' || trimmed[start] == '-')))
                {
                    return false;
                }

                string token = trimmed.Substring(start, index - start);
                if (!int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out int decValue))
                {
                    return false;
                }

                if (decValue < 0 || decValue > 255)
                {
                    return false;
                }

                ch = (char)decValue;
                isTerminator = decValue == 0;
                return true;
            }
        }

        private static bool TryExtractQuotedStringContent(string value, out string content)
        {
            content = string.Empty;
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            int start = value.IndexOf('"');
            if (start < 0 || start >= value.Length - 1)
            {
                return false;
            }

            int i = start + 1;
            bool escaped = false;
            while (i < value.Length)
            {
                char ch = value[i];
                if (escaped)
                {
                    escaped = false;
                    i++;
                    continue;
                }

                if (ch == '\\')
                {
                    escaped = true;
                    i++;
                    continue;
                }

                if (ch == '"')
                {
                    content = value.Substring(start + 1, i - start - 1);
                    return true;
                }

                i++;
            }

            return false;
        }

        private static bool TryExtractCharsFromSingleQuotedItems(string value, out string content)
        {
            content = string.Empty;
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            var chars = new List<char>(4);
            int i = 0;
            while (i < value.Length && chars.Count < 4)
            {
                if (value[i] != '\'')
                {
                    i++;
                    continue;
                }

                i++;
                if (i >= value.Length)
                {
                    break;
                }

                char parsed;
                if (value[i] == '\\')
                {
                    i++;
                    if (i >= value.Length)
                    {
                        break;
                    }

                    parsed = DecodeEscape(value[i]);
                    i++;
                }
                else
                {
                    parsed = value[i];
                    i++;
                }

                if (i >= value.Length || value[i] != '\'')
                {
                    continue;
                }

                chars.Add(parsed);
                i++;
            }

            if (chars.Count == 0)
            {
                return false;
            }

            content = new string(chars.ToArray());
            return true;
        }

        private static char DecodeEscape(char ch)
        {
            switch (ch)
            {
                case 'n':
                    return '\n';
                case 'r':
                    return '\r';
                case 't':
                    return '\t';
                case '\\':
                    return '\\';
                case '\'':
                    return '\'';
                case '"':
                    return '"';
                case '0':
                    return '\0';
                default:
                    return ch;
            }
        }

        private static bool IsFunctionPointerType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            string compact = RemoveWhitespace(type).ToLowerInvariant();
            return compact.Contains("(*)(") ||
                   compact.Contains("(__cdecl*)(") ||
                   compact.Contains("(__stdcall*)(") ||
                   compact.Contains("(__fastcall*)(") ||
                   compact.Contains("(__thiscall*)(") ||
                   compact.Contains("function");
        }

        private static bool IsStructValueType(string type, string normalizedValue)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            if (type.IndexOf('*') >= 0 || type.IndexOf('&') >= 0)
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            if (ContainsTypeWord(lower, "struct"))
            {
                return true;
            }

            if (!IsBraceLikeValue(normalizedValue))
            {
                return false;
            }

            if (type.IndexOf('[') >= 0 || type.IndexOf(']') >= 0)
            {
                return false;
            }

            return IsLikelyUserDefinedType(type);
        }

        private static bool IsLikelyClassOrStructPointerType(string type)
        {
            if (string.IsNullOrWhiteSpace(type) || type.IndexOf('*') < 0)
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            if (lower.Contains(" class ") || lower.StartsWith("class ", StringComparison.Ordinal) ||
                lower.Contains(" struct ") || lower.StartsWith("struct ", StringComparison.Ordinal))
            {
                return true;
            }

            string normalized = lower
                .Replace('*', ' ')
                .Replace('&', ' ')
                .Replace('(', ' ')
                .Replace(')', ' ')
                .Replace('[', ' ')
                .Replace(']', ' ')
                .Replace(',', ' ')
                .Replace('<', ' ')
                .Replace('>', ' ')
                .Replace(':', ' ');

            string[] tokens = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string token in tokens)
            {
                if (TypeNoiseTokens.Contains(token))
                {
                    continue;
                }

                if (PrimitiveTypeTokens.Contains(token))
                {
                    return false;
                }

                return true;
            }

            return false;
        }

        private static bool IsLikelyUserDefinedType(string type)
        {
            string lower = type.ToLowerInvariant();
            string normalized = lower
                .Replace('*', ' ')
                .Replace('&', ' ')
                .Replace('(', ' ')
                .Replace(')', ' ')
                .Replace('[', ' ')
                .Replace(']', ' ')
                .Replace(',', ' ')
                .Replace('<', ' ')
                .Replace('>', ' ')
                .Replace(':', ' ');

            string[] tokens = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string token in tokens)
            {
                if (TypeNoiseTokens.Contains(token))
                {
                    continue;
                }

                if (PrimitiveTypeTokens.Contains(token))
                {
                    return false;
                }

                if (string.Equals(token, "std", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                return true;
            }

            return false;
        }

        private static string ExtractTrailingIdentifier(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return string.Empty;
            }

            int end = name.Length - 1;
            while (end >= 0 && !IsWordChar(name[end]))
            {
                end--;
            }

            if (end < 0)
            {
                return string.Empty;
            }

            int start = end;
            while (start >= 0 && IsWordChar(name[start]))
            {
                start--;
            }

            return name.Substring(start + 1, end - start);
        }

        private static bool IsAddressLikeValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            int i = 0;
            while (i < value.Length && char.IsWhiteSpace(value[i]))
            {
                i++;
            }

            if (i + 3 >= value.Length || value[i] != '0' || (value[i + 1] != 'x' && value[i + 1] != 'X'))
            {
                return false;
            }

            int hexCount = 0;
            i += 2;
            while (i < value.Length)
            {
                char ch = value[i];
                bool isHex =
                    (ch >= '0' && ch <= '9') ||
                    (ch >= 'a' && ch <= 'f') ||
                    (ch >= 'A' && ch <= 'F');
                if (!isHex)
                {
                    break;
                }

                hexCount++;
                i++;
            }

            return hexCount >= 4;
        }

        private static bool IsBraceLikeValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            int i = 0;
            while (i < value.Length && char.IsWhiteSpace(value[i]))
            {
                i++;
            }

            return i < value.Length && value[i] == '{';
        }

        private static bool ContainsTypeWord(string typeLower, string wordLower)
        {
            int index = 0;
            while (index < typeLower.Length)
            {
                index = typeLower.IndexOf(wordLower, index, StringComparison.Ordinal);
                if (index < 0)
                {
                    return false;
                }

                bool leftOk = index == 0 || !IsWordChar(typeLower[index - 1]);
                int rightIndex = index + wordLower.Length;
                bool rightOk = rightIndex >= typeLower.Length || !IsWordChar(typeLower[rightIndex]);
                if (leftOk && rightOk)
                {
                    return true;
                }

                index = rightIndex;
            }

            return false;
        }

        private static bool IsWordChar(char ch)
        {
            return ch == '_' || char.IsLetterOrDigit(ch);
        }

        private static bool IsValidIdentifier(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (!(value[0] == '_' || char.IsLetter(value[0])))
            {
                return false;
            }

            for (int i = 1; i < value.Length; i++)
            {
                if (!IsWordChar(value[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private static HashSet<string> GetStackFrameParameterNames(StackFrame2 stackFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var names = new HashSet<string>(StringComparer.Ordinal);
            if (stackFrame == null)
            {
                return names;
            }

            Expressions arguments;
            try
            {
                arguments = stackFrame.Arguments2[true];
            }
            catch (COMException)
            {
                return names;
            }

            if (arguments == null)
            {
                return names;
            }

            foreach (EnvDTE.Expression argument in arguments)
            {
                if (argument == null || string.IsNullOrWhiteSpace(argument.Name))
                {
                    continue;
                }

                string identifier = ExtractTrailingIdentifier(argument.Name);
                if (!IsIdentifierName(identifier) || string.Equals(identifier, "this", StringComparison.Ordinal))
                {
                    continue;
                }

                names.Add(identifier);
            }

            return names;
        }

        private static int AddMissingParameterAnchorTags(
            ITextSnapshot snapshot,
            int functionStartLineNumber,
            HashSet<string> parameterNames,
            Dictionary<string, string> valueMap,
            Dictionary<string, int> highlightLevels,
            InlineValueDisplayMode displayMode,
            Dictionary<string, TagEntry> firstEntryByIdentifier,
            Dictionary<string, TagEntry> latestEntryByIdentifier)
        {
            if (snapshot == null ||
                functionStartLineNumber < 0 ||
                parameterNames == null ||
                parameterNames.Count == 0 ||
                valueMap == null ||
                valueMap.Count == 0)
            {
                return int.MaxValue;
            }

            int signatureStartLine = FindFunctionSignatureStartLine(snapshot, functionStartLineNumber);
            Dictionary<string, InsertionAnchor> anchors = FindParameterAnchors(snapshot, signatureStartLine, functionStartLineNumber, parameterNames, displayMode);
            if (anchors.Count == 0)
            {
                return int.MaxValue;
            }

            int earliestAddedLine = int.MaxValue;
            foreach (string parameterName in parameterNames)
            {
                if (latestEntryByIdentifier.ContainsKey(parameterName))
                {
                    continue;
                }

                if (!valueMap.TryGetValue(parameterName, out string value))
                {
                    continue;
                }

                if (!anchors.TryGetValue(parameterName, out InsertionAnchor anchor))
                {
                    continue;
                }

                int highlightLevel = 0;
                if (highlightLevels.TryGetValue(parameterName, out int level))
                {
                    highlightLevel = level;
                }

                ValueDisplayKind displayKind = GetDisplayKind(value);
                string cleanedValue = StripDisplayMarker(value);

                string displayText = displayMode == InlineValueDisplayMode.EndOfLine
                    ? parameterName + ": " + cleanedValue
                    : cleanedValue;

                var entry = new TagEntry(anchor.Position, displayText, highlightLevel, anchor.Affinity, displayKind);
                latestEntryByIdentifier[parameterName] = entry;
                if (!firstEntryByIdentifier.ContainsKey(parameterName))
                {
                    firstEntryByIdentifier[parameterName] = entry;
                }

                int lineNumber = snapshot.GetLineNumberFromPosition(anchor.Position);
                if (lineNumber < earliestAddedLine)
                {
                    earliestAddedLine = lineNumber;
                }
            }

            return earliestAddedLine;
        }

        private static Dictionary<string, InsertionAnchor> FindParameterAnchors(
            ITextSnapshot snapshot,
            int signatureStartLine,
            int functionStartLine,
            HashSet<string> parameterNames,
            InlineValueDisplayMode displayMode)
        {
            var anchors = new Dictionary<string, InsertionAnchor>(StringComparer.Ordinal);
            if (snapshot == null || parameterNames == null || parameterNames.Count == 0)
            {
                return anchors;
            }

            int startLine = Math.Max(0, signatureStartLine);
            int endLine = Math.Min(functionStartLine, snapshot.LineCount - 1);
            if (startLine > endLine)
            {
                return anchors;
            }

            for (int lineIndex = startLine; lineIndex <= endLine; lineIndex++)
            {
                ITextSnapshotLine line = snapshot.GetLineFromLineNumber(lineIndex);
                string lineText = line.GetText();
                List<CppCurrentLineTokenizer.IdentifierToken> tokens = CppCurrentLineTokenizer.TokenizeIdentifiers(lineText);
                foreach (CppCurrentLineTokenizer.IdentifierToken token in tokens)
                {
                    if (!parameterNames.Contains(token.Name) || anchors.ContainsKey(token.Name))
                    {
                        continue;
                    }

                    if (IsLikelyFunctionInvocation(lineText, token))
                    {
                        continue;
                    }

                    int position;
                    PositionAffinity affinity;
                    if (displayMode == InlineValueDisplayMode.EndOfLine)
                    {
                        position = line.End.Position;
                        affinity = PositionAffinity.Predecessor;
                    }
                    else
                    {
                        position = line.Start.Position + token.Start + token.Length;
                        affinity = PositionAffinity.Successor;
                    }

                    anchors[token.Name] = new InsertionAnchor(position, affinity);
                }
            }

            return anchors;
        }

        private static int FindFunctionSignatureStartLine(ITextSnapshot snapshot, int functionStartLine)
        {
            if (snapshot == null || functionStartLine < 0 || functionStartLine >= snapshot.LineCount)
            {
                return functionStartLine;
            }

            int scanStart = Math.Max(0, functionStartLine - 40);
            int parenDepth = 0;
            for (int lineIndex = functionStartLine; lineIndex >= scanStart; lineIndex--)
            {
                string text = snapshot.GetLineFromLineNumber(lineIndex).GetText();
                for (int i = text.Length - 1; i >= 0; i--)
                {
                    char ch = text[i];
                    if (ch == ')')
                    {
                        parenDepth++;
                    }
                    else if (ch == '(')
                    {
                        if (parenDepth == 0)
                        {
                            return lineIndex;
                        }

                        parenDepth--;
                        if (parenDepth == 0)
                        {
                            // We just matched the outermost ')' for the function parameter list.
                            return lineIndex;
                        }
                    }
                    else if (ch == ';' && parenDepth == 0)
                    {
                        return functionStartLine;
                    }
                }
            }

            return functionStartLine;
        }

        private static int FindEnclosingFunctionStartLine(ITextSnapshot snapshot, int targetLineNumber)
        {
            if (snapshot == null || targetLineNumber < 0 || snapshot.LineCount == 0)
            {
                return -1;
            }

            int lastLine = Math.Min(targetLineNumber, snapshot.LineCount - 1);
            bool inBlockComment = false;
            var blockStack = new List<BlockScope>();
            for (int lineIndex = 0; lineIndex <= lastLine; lineIndex++)
            {
                string text = snapshot.GetLineFromLineNumber(lineIndex).GetText();
                bool inQuotedLiteral = false;
                char quoteDelimiter = '\0';
                bool escaped = false;
                for (int i = 0; i < text.Length; i++)
                {
                    char ch = text[i];
                    char next = i + 1 < text.Length ? text[i + 1] : '\0';

                    if (inBlockComment)
                    {
                        if (ch == '*' && next == '/')
                        {
                            inBlockComment = false;
                            i++;
                        }

                        continue;
                    }

                    if (inQuotedLiteral)
                    {
                        if (escaped)
                        {
                            escaped = false;
                            continue;
                        }

                        if (ch == '\\')
                        {
                            escaped = true;
                            continue;
                        }

                        if (ch == quoteDelimiter)
                        {
                            inQuotedLiteral = false;
                        }

                        continue;
                    }

                    if (ch == '/' && next == '/')
                    {
                        break;
                    }

                    if (ch == '/' && next == '*')
                    {
                        inBlockComment = true;
                        i++;
                        continue;
                    }

                    if (ch == '"' || ch == '\'')
                    {
                        inQuotedLiteral = true;
                        quoteDelimiter = ch;
                        escaped = false;
                        continue;
                    }

                    if (ch == '{')
                    {
                        bool isFunctionStart = IsLikelyFunctionOpeningBrace(snapshot, lineIndex, i);
                        blockStack.Add(new BlockScope(lineIndex, isFunctionStart));
                    }
                    else if (ch == '}')
                    {
                        if (blockStack.Count > 0)
                        {
                            blockStack.RemoveAt(blockStack.Count - 1);
                        }
                    }
                }
            }

            for (int i = blockStack.Count - 1; i >= 0; i--)
            {
                BlockScope scope = blockStack[i];
                if (scope.IsFunctionStart)
                {
                    return scope.LineNumber;
                }
            }

            return -1;
        }

        private static int FindEnclosingFunctionEndLine(ITextSnapshot snapshot, int functionStartLine)
        {
            if (snapshot == null || functionStartLine < 0 || functionStartLine >= snapshot.LineCount)
            {
                return functionStartLine;
            }

            bool inBlockComment = false;
            bool seenFunctionOpeningBrace = false;
            int braceDepth = 0;
            for (int lineIndex = functionStartLine; lineIndex < snapshot.LineCount; lineIndex++)
            {
                string text = snapshot.GetLineFromLineNumber(lineIndex).GetText();
                bool inQuotedLiteral = false;
                char quoteDelimiter = '\0';
                bool escaped = false;
                for (int i = 0; i < text.Length; i++)
                {
                    char ch = text[i];
                    char next = i + 1 < text.Length ? text[i + 1] : '\0';

                    if (inBlockComment)
                    {
                        if (ch == '*' && next == '/')
                        {
                            inBlockComment = false;
                            i++;
                        }

                        continue;
                    }

                    if (inQuotedLiteral)
                    {
                        if (escaped)
                        {
                            escaped = false;
                            continue;
                        }

                        if (ch == '\\')
                        {
                            escaped = true;
                            continue;
                        }

                        if (ch == quoteDelimiter)
                        {
                            inQuotedLiteral = false;
                        }

                        continue;
                    }

                    if (ch == '/' && next == '/')
                    {
                        break;
                    }

                    if (ch == '/' && next == '*')
                    {
                        inBlockComment = true;
                        i++;
                        continue;
                    }

                    if (ch == '"' || ch == '\'')
                    {
                        inQuotedLiteral = true;
                        quoteDelimiter = ch;
                        escaped = false;
                        continue;
                    }

                    if (ch == '{')
                    {
                        if (!seenFunctionOpeningBrace)
                        {
                            seenFunctionOpeningBrace = true;
                        }

                        braceDepth++;
                    }
                    else if (ch == '}')
                    {
                        if (!seenFunctionOpeningBrace)
                        {
                            continue;
                        }

                        braceDepth--;
                        if (braceDepth <= 0)
                        {
                            return lineIndex;
                        }
                    }
                }
            }

            return snapshot.LineCount - 1;
        }

        private static bool IsLikelyFunctionOpeningBrace(ITextSnapshot snapshot, int braceLineNumber, int braceColumn)
        {
            int startLine = Math.Max(0, braceLineNumber - 8);
            var builder = new StringBuilder(256);
            for (int lineIndex = startLine; lineIndex <= braceLineNumber; lineIndex++)
            {
                string lineText = snapshot.GetLineFromLineNumber(lineIndex).GetText();
                if (lineIndex == braceLineNumber)
                {
                    int length = Math.Min(braceColumn, lineText.Length);
                    lineText = lineText.Substring(0, length);
                }

                builder.Append(lineText);
                builder.Append(' ');
            }

            string compact = RemoveWhitespace(builder.ToString());
            if (compact.Length == 0)
            {
                return false;
            }

            int closeParen = compact.LastIndexOf(')');
            if (closeParen < 0)
            {
                return false;
            }

            int openParen = compact.LastIndexOf('(', closeParen);
            if (openParen < 0)
            {
                return false;
            }

            string beforeParen = compact.Substring(0, openParen);
            if (beforeParen.Length == 0)
            {
                return false;
            }

            string trailing = ExtractTrailingIdentifier(beforeParen);
            if (NonFunctionBlockKeywords.Contains(trailing))
            {
                return false;
            }

            if (beforeParen.EndsWith("=", StringComparison.Ordinal) ||
                beforeParen.EndsWith("return", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return true;
        }

        private static bool IsIdentifierName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            if (!(name[0] == '_' || char.IsLetter(name[0])))
            {
                return false;
            }

            for (int i = 1; i < name.Length; i++)
            {
                if (!IsWordChar(name[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private static ValueDisplayKind GetDisplayKind(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return ValueDisplayKind.Normal;
            }

            if (value.StartsWith(NullValueMarker, StringComparison.Ordinal))
            {
                return ValueDisplayKind.Null;
            }

            if (value.StartsWith(UninitializedValueMarker, StringComparison.Ordinal))
            {
                return ValueDisplayKind.Uninitialized;
            }

            return ValueDisplayKind.Normal;
        }

        private static string StripDisplayMarker(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }

            if (value.StartsWith(NullValueMarker, StringComparison.Ordinal))
            {
                return value.Substring(NullValueMarker.Length);
            }

            if (value.StartsWith(UninitializedValueMarker, StringComparison.Ordinal))
            {
                return value.Substring(UninitializedValueMarker.Length);
            }

            return value;
        }

        private static bool TryParseGetterCall(string lineText, CppCurrentLineTokenizer.IdentifierToken token, out GetterCallToken getterCall)
        {
            getterCall = default;
            if (!IsGetterLikeName(token.Name))
            {
                return false;
            }

            int openParenIndex = token.Start + token.Length;
            while (openParenIndex < lineText.Length && char.IsWhiteSpace(lineText[openParenIndex]))
            {
                openParenIndex++;
            }

            if (openParenIndex >= lineText.Length || lineText[openParenIndex] != '(')
            {
                return false;
            }

            if (!TryFindMatchingParen(lineText, openParenIndex, out int closeParenIndex))
            {
                return false;
            }

            string argumentsText = lineText.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1);
            if (!string.IsNullOrWhiteSpace(argumentsText))
            {
                return false;
            }

            int expressionStart = FindGetterExpressionStart(lineText, token.Start);
            if (expressionStart < 0 || expressionStart > token.Start || closeParenIndex < expressionStart)
            {
                return false;
            }

            string expressionText = RemoveWhitespace(lineText.Substring(expressionStart, closeParenIndex - expressionStart + 1).Trim());
            if (string.IsNullOrEmpty(expressionText))
            {
                return false;
            }

            getterCall = new GetterCallToken(
                expressionText,
                expressionText,
                expressionText,
                token.Name,
                closeParenIndex + 1,
                InlineValueEvaluationKinds.GetterCalls);
            return true;
        }

        private static bool TryParseArrayIndexAccess(string lineText, CppCurrentLineTokenizer.IdentifierToken token, out GetterCallToken indexedCall)
        {
            indexedCall = default;
            if (string.IsNullOrEmpty(lineText))
            {
                return false;
            }

            int openBracketIndex = token.Start + token.Length;
            while (openBracketIndex < lineText.Length && char.IsWhiteSpace(lineText[openBracketIndex]))
            {
                openBracketIndex++;
            }

            if (openBracketIndex >= lineText.Length || lineText[openBracketIndex] != '[')
            {
                return false;
            }

            if (!TryFindMatchingBracket(lineText, openBracketIndex, out int closeBracketIndex))
            {
                return false;
            }

            string indexText = lineText.Substring(openBracketIndex + 1, closeBracketIndex - openBracketIndex - 1).Trim();
            if (string.IsNullOrEmpty(indexText))
            {
                return false;
            }

            if (ContainsIndexedExpressionSideEffects(indexText))
            {
                return false;
            }

            string expressionText = RemoveWhitespace(lineText.Substring(token.Start, closeBracketIndex - token.Start + 1).Trim());
            if (string.IsNullOrEmpty(expressionText))
            {
                return false;
            }

            indexedCall = new GetterCallToken(
                expressionText,
                expressionText,
                expressionText,
                token.Name,
                closeBracketIndex + 1,
                InlineValueEvaluationKinds.IndexedExpressions);
            return true;
        }

        private static bool TryParseMemberAccessExpression(string lineText, CppCurrentLineTokenizer.IdentifierToken token, out GetterCallToken memberAccess)
        {
            memberAccess = default;
            if (string.IsNullOrEmpty(lineText))
            {
                return false;
            }

            if (!HasMemberAccessOperatorBeforeToken(lineText, token.Start) ||
                HasMemberAccessContinuationAfterToken(lineText, token.Start + token.Length))
            {
                return false;
            }

            if (IsLikelyFunctionInvocation(lineText, token))
            {
                return false;
            }

            int expressionStart = FindGetterExpressionStart(lineText, token.Start);
            if (expressionStart < 0 || expressionStart >= token.Start)
            {
                return false;
            }

            string expressionText = RemoveWhitespace(lineText.Substring(expressionStart, token.Start + token.Length - expressionStart).Trim());
            if (string.IsNullOrEmpty(expressionText) || ContainsIndexedExpressionSideEffects(expressionText))
            {
                return false;
            }

            memberAccess = new GetterCallToken(
                expressionText,
                expressionText,
                expressionText,
                token.Name,
                token.Start + token.Length,
                InlineValueEvaluationKinds.FallbackExpressions);
            return true;
        }

        private static bool HasMemberAccessOperatorBeforeToken(string lineText, int tokenStart)
        {
            if (string.IsNullOrEmpty(lineText) || tokenStart <= 0)
            {
                return false;
            }

            int index = SkipWhitespaceBackward(lineText, tokenStart - 1);
            if (index < 0)
            {
                return false;
            }

            if (lineText[index] == '.')
            {
                return true;
            }

            return (index >= 1 && lineText[index] == '>' && lineText[index - 1] == '-') ||
                   (index >= 1 && lineText[index] == ':' && lineText[index - 1] == ':');
        }

        private static bool HasMemberAccessContinuationAfterToken(string lineText, int tokenEnd)
        {
            if (string.IsNullOrEmpty(lineText) || tokenEnd < 0)
            {
                return false;
            }

            int index = tokenEnd;
            while (index < lineText.Length && char.IsWhiteSpace(lineText[index]))
            {
                index++;
            }

            if (index >= lineText.Length)
            {
                return false;
            }

            if (lineText[index] == '.')
            {
                return true;
            }

            return (index + 1 < lineText.Length && lineText[index] == '-' && lineText[index + 1] == '>') ||
                   (index + 1 < lineText.Length && lineText[index] == ':' && lineText[index + 1] == ':');
        }

        private static bool ContainsIndexedExpressionSideEffects(string expressionText)
        {
            if (string.IsNullOrWhiteSpace(expressionText))
            {
                return false;
            }

            for (int i = 0; i < expressionText.Length; i++)
            {
                char ch = expressionText[i];
                if (ch == '"' || ch == '\'')
                {
                    i = SkipQuotedLiteral(expressionText, i) - 1;
                    continue;
                }

                if (ch == '+' || ch == '-')
                {
                    if (i + 1 < expressionText.Length && expressionText[i + 1] == ch)
                    {
                        return true;
                    }

                    continue;
                }

                if (ch == '=')
                {
                    char prev = i > 0 ? expressionText[i - 1] : '\0';
                    char next = i + 1 < expressionText.Length ? expressionText[i + 1] : '\0';
                    bool isComparison =
                        prev == '=' ||
                        prev == '!' ||
                        prev == '<' ||
                        prev == '>' ||
                        next == '=';
                    if (!isComparison)
                    {
                        return true;
                    }

                    continue;
                }
            }

            return false;
        }

        private static bool TryFindMatchingParen(string lineText, int openParenIndex, out int closeParenIndex)
        {
            closeParenIndex = -1;
            int depth = 0;
            int i = openParenIndex;
            while (i < lineText.Length)
            {
                char ch = lineText[i];
                if (ch == '/' && i + 1 < lineText.Length)
                {
                    char next = lineText[i + 1];
                    if (next == '/' || next == '*')
                    {
                        break;
                    }
                }

                if (ch == '"' || ch == '\'')
                {
                    i = SkipQuotedLiteral(lineText, i);
                    continue;
                }

                if (ch == '(')
                {
                    depth++;
                }
                else if (ch == ')')
                {
                    depth--;
                    if (depth == 0)
                    {
                        closeParenIndex = i;
                        return true;
                    }
                }

                i++;
            }

            return false;
        }

        private static int SkipQuotedLiteral(string line, int start)
        {
            char quote = line[start];
            int i = start + 1;
            while (i < line.Length)
            {
                if (line[i] == '\\')
                {
                    i += 2;
                    continue;
                }

                if (line[i] == quote)
                {
                    return i + 1;
                }

                i++;
            }

            return line.Length;
        }

        private static int FindGetterExpressionStart(string lineText, int tokenStart)
        {
            int start = tokenStart;
            while (true)
            {
                int i = start - 1;
                while (i >= 0 && char.IsWhiteSpace(lineText[i]))
                {
                    i--;
                }

                if (i < 0)
                {
                    break;
                }

                int opStart;
                if (lineText[i] == '.')
                {
                    opStart = i;
                }
                else if (i >= 1 && lineText[i] == '>' && lineText[i - 1] == '-')
                {
                    opStart = i - 1;
                }
                else if (i >= 1 && lineText[i] == ':' && lineText[i - 1] == ':')
                {
                    opStart = i - 1;
                }
                else
                {
                    break;
                }

                int leftEnd = opStart - 1;
                while (leftEnd >= 0 && char.IsWhiteSpace(lineText[leftEnd]))
                {
                    leftEnd--;
                }

                if (!TryFindReceiverExpressionStart(lineText, leftEnd, out int leftStart))
                {
                    break;
                }

                start = leftStart;
            }

            return start;
        }

        private static bool TryFindReceiverExpressionStart(string lineText, int receiverEnd, out int receiverStart)
        {
            receiverStart = -1;
            int index = SkipWhitespaceBackward(lineText, receiverEnd);
            if (index < 0)
            {
                return false;
            }

            while (index >= 0)
            {
                char ch = lineText[index];
                if (IsWordChar(ch))
                {
                    int wordStart = index;
                    while (wordStart >= 0 && IsWordChar(lineText[wordStart]))
                    {
                        wordStart--;
                    }

                    receiverStart = wordStart + 1;
                    index = SkipWhitespaceBackward(lineText, receiverStart - 1);
                    continue;
                }

                if (ch == ']')
                {
                    if (!TryFindMatchingPairBackward(lineText, index, '[', ']', out int openBracketIndex))
                    {
                        return false;
                    }

                    receiverStart = openBracketIndex;
                    index = SkipWhitespaceBackward(lineText, openBracketIndex - 1);
                    continue;
                }

                if (ch == ')')
                {
                    if (!TryFindMatchingPairBackward(lineText, index, '(', ')', out int openParenIndex))
                    {
                        return false;
                    }

                    int beforeParen = SkipWhitespaceBackward(lineText, openParenIndex - 1);
                    if (beforeParen >= 0 &&
                        (IsWordChar(lineText[beforeParen]) || lineText[beforeParen] == ']' || lineText[beforeParen] == ')'))
                    {
                        return false;
                    }

                    receiverStart = openParenIndex;
                    index = SkipWhitespaceBackward(lineText, openParenIndex - 1);
                    continue;
                }

                break;
            }

            return receiverStart >= 0;
        }

        private static bool TryFindMatchingPairBackward(string text, int closeIndex, char openChar, char closeChar, out int openIndex)
        {
            openIndex = -1;
            if (string.IsNullOrEmpty(text) || closeIndex < 0 || closeIndex >= text.Length || text[closeIndex] != closeChar)
            {
                return false;
            }

            int depth = 0;
            for (int index = closeIndex; index >= 0; index--)
            {
                char ch = text[index];
                if (ch == closeChar)
                {
                    depth++;
                }
                else if (ch == openChar)
                {
                    depth--;
                    if (depth == 0)
                    {
                        openIndex = index;
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsLikelyFunctionInvocation(string lineText, CppCurrentLineTokenizer.IdentifierToken token)
        {
            int i = token.Start + token.Length;
            while (i < lineText.Length && char.IsWhiteSpace(lineText[i]))
            {
                i++;
            }

            return i < lineText.Length && lineText[i] == '(';
        }

        private static bool IsPointerReceiverForGetterCall(string lineText, CppCurrentLineTokenizer.IdentifierToken token)
        {
            if (string.IsNullOrEmpty(lineText))
            {
                return false;
            }

            int i = token.Start + token.Length;
            while (i < lineText.Length && char.IsWhiteSpace(lineText[i]))
            {
                i++;
            }

            if (i + 1 >= lineText.Length || lineText[i] != '-' || lineText[i + 1] != '>')
            {
                return false;
            }

            i += 2;
            while (i < lineText.Length && char.IsWhiteSpace(lineText[i]))
            {
                i++;
            }

            if (i >= lineText.Length || !(lineText[i] == '_' || char.IsLetter(lineText[i])))
            {
                return false;
            }

            int nameStart = i;
            i++;
            while (i < lineText.Length && IsWordChar(lineText[i]))
            {
                i++;
            }

            string memberName = lineText.Substring(nameStart, i - nameStart);
            if (!IsGetterLikeName(memberName))
            {
                return false;
            }

            while (i < lineText.Length && char.IsWhiteSpace(lineText[i]))
            {
                i++;
            }

            return i < lineText.Length && lineText[i] == '(';
        }

        private static bool IsGetterLikeName(string methodName)
        {
            if (string.IsNullOrWhiteSpace(methodName))
            {
                return false;
            }

            return methodName.StartsWith("get", StringComparison.OrdinalIgnoreCase) ||
                   methodName.StartsWith("is", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryBindVisibleDirectReturnGetter(ITextSnapshot snapshot, GetterCallToken getterCall, out GetterCallToken boundGetterCall)
        {
            boundGetterCall = getterCall;
            if (!TryGetVisibleDirectReturnGetterExpression(snapshot, getterCall.MethodName, out string returnExpression))
            {
                return false;
            }

            if (!TryBuildGetterEvaluationExpression(getterCall.ExpressionText, getterCall.MethodName, returnExpression, out string evaluationExpressionText))
            {
                return false;
            }

            boundGetterCall = getterCall.WithEvaluationExpression(evaluationExpressionText);
            return true;
        }

        private static bool TryGetVisibleDirectReturnGetterExpression(ITextSnapshot snapshot, string methodName, out string returnExpression)
        {
            returnExpression = null;
            if (snapshot == null || !IsValidIdentifier(methodName))
            {
                return false;
            }

            string text = snapshot.GetText();
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            int searchIndex = 0;
            while (searchIndex < text.Length)
            {
                int methodIndex = text.IndexOf(methodName, searchIndex, StringComparison.Ordinal);
                if (methodIndex < 0)
                {
                    break;
                }

                searchIndex = methodIndex + methodName.Length;
                int methodEnd = methodIndex + methodName.Length;
                bool leftBoundary = methodIndex == 0 || !IsWordChar(text[methodIndex - 1]);
                bool rightBoundary = methodEnd >= text.Length || !IsWordChar(text[methodEnd]);
                if (!leftBoundary || !rightBoundary)
                {
                    continue;
                }

                int openParenIndex = SkipWhitespace(text, methodEnd);
                if (openParenIndex >= text.Length || text[openParenIndex] != '(')
                {
                    continue;
                }

                if (!TryFindMatchingPair(text, openParenIndex, '(', ')', out int closeParenIndex))
                {
                    continue;
                }

                string argumentsText = text.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1);
                if (!IsZeroArgumentSignature(argumentsText))
                {
                    continue;
                }

                int bodyStart = TryFindVisibleGetterBodyStart(text, closeParenIndex + 1);
                if (bodyStart < 0 || bodyStart >= text.Length || text[bodyStart] != '{')
                {
                    continue;
                }

                if (!TryFindMatchingPair(text, bodyStart, '{', '}', out int closeBraceIndex))
                {
                    continue;
                }

                string body = text.Substring(bodyStart + 1, closeBraceIndex - bodyStart - 1);
                if (TryExtractDirectReturnGetterExpression(body, out returnExpression))
                {
                    return true;
                }

                searchIndex = closeBraceIndex + 1;
            }

            return false;
        }

        private static bool IsZeroArgumentSignature(string argumentsText)
        {
            if (string.IsNullOrWhiteSpace(argumentsText))
            {
                return true;
            }

            string compact = RemoveWhitespace(StripComments(argumentsText));
            return string.Equals(compact, "void", StringComparison.Ordinal);
        }

        private static bool TryBuildGetterEvaluationExpression(string getterExpression, string methodName, string returnExpression, out string evaluationExpression)
        {
            evaluationExpression = null;
            if (string.IsNullOrWhiteSpace(getterExpression) || string.IsNullOrWhiteSpace(returnExpression) || string.IsNullOrWhiteSpace(methodName))
            {
                return false;
            }

            string compactReturnExpression = RemoveWhitespace(returnExpression.Trim());
            if (string.IsNullOrEmpty(compactReturnExpression))
            {
                return false;
            }

            int callSuffixIndex = getterExpression.LastIndexOf(methodName + "()", StringComparison.Ordinal);
            if (callSuffixIndex < 0)
            {
                return false;
            }

            int methodStart = callSuffixIndex;
            if (methodStart >= 2 && getterExpression[methodStart - 1] == '>' && getterExpression[methodStart - 2] == '-')
            {
                string receiver = getterExpression.Substring(0, methodStart - 2);
                evaluationExpression = BuildGetterReturnExpression(receiver, "->", compactReturnExpression);
                return !string.IsNullOrEmpty(evaluationExpression);
            }

            if (methodStart >= 1 && getterExpression[methodStart - 1] == '.')
            {
                string receiver = getterExpression.Substring(0, methodStart - 1);
                evaluationExpression = BuildGetterReturnExpression(receiver, ".", compactReturnExpression);
                return !string.IsNullOrEmpty(evaluationExpression);
            }

            evaluationExpression = compactReturnExpression;
            return true;
        }

        private static string BuildGetterReturnExpression(string receiver, string receiverOperator, string returnExpression)
        {
            if (string.IsNullOrWhiteSpace(returnExpression))
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(receiver))
            {
                return returnExpression;
            }

            if (returnExpression.StartsWith("this->", StringComparison.Ordinal))
            {
                string memberSuffix = returnExpression.Substring("this->".Length);
                return string.IsNullOrEmpty(memberSuffix) ? null : receiver + receiverOperator + memberSuffix;
            }

            if (returnExpression.StartsWith("this.", StringComparison.Ordinal))
            {
                string memberSuffix = returnExpression.Substring("this.".Length);
                return string.IsNullOrEmpty(memberSuffix) ? null : receiver + receiverOperator + memberSuffix;
            }

            if (returnExpression.StartsWith("::", StringComparison.Ordinal) || returnExpression.Contains("::"))
            {
                return returnExpression;
            }

            return receiver + receiverOperator + returnExpression;
        }

        private static bool TryFindMatchingBracket(string lineText, int openBracketIndex, out int closeBracketIndex)
        {
            closeBracketIndex = -1;
            int depth = 0;
            int i = openBracketIndex;
            while (i < lineText.Length)
            {
                char ch = lineText[i];
                if (ch == '/' && i + 1 < lineText.Length)
                {
                    char next = lineText[i + 1];
                    if (next == '/' || next == '*')
                    {
                        break;
                    }
                }

                if (ch == '"' || ch == '\'')
                {
                    i = SkipQuotedLiteral(lineText, i);
                    continue;
                }

                if (ch == '[')
                {
                    depth++;
                }
                else if (ch == ']')
                {
                    depth--;
                    if (depth == 0)
                    {
                        closeBracketIndex = i;
                        return true;
                    }
                }

                i++;
            }

            return false;
        }

        private static int TryFindVisibleGetterBodyStart(string text, int startIndex)
        {
            int index = SkipWhitespaceAndComments(text, startIndex);
            while (index >= 0 && index < text.Length)
            {
                if (text[index] == '{')
                {
                    return index;
                }

                if (text[index] == ';')
                {
                    return -1;
                }

                if (text[index] == '-' && index + 1 < text.Length && text[index + 1] == '>')
                {
                    index += 2;
                    while (index < text.Length)
                    {
                        if (text[index] == '{')
                        {
                            return index;
                        }

                        if (text[index] == ';')
                        {
                            return -1;
                        }

                        if (text[index] == '\r' || text[index] == '\n')
                        {
                            break;
                        }

                        index++;
                    }
                }
                else
                {
                    int tokenStart = index;
                    index = ReadWordToken(text, index);
                    if (tokenStart == index)
                    {
                        return -1;
                    }

                    string token = text.Substring(tokenStart, index - tokenStart);
                    if (!string.Equals(token, "const", StringComparison.Ordinal) &&
                        !string.Equals(token, "noexcept", StringComparison.Ordinal) &&
                        !string.Equals(token, "override", StringComparison.Ordinal) &&
                        !string.Equals(token, "final", StringComparison.Ordinal))
                    {
                        return -1;
                    }
                }

                index = SkipWhitespaceAndComments(text, index);
            }

            return -1;
        }

        private static bool TryFindMatchingPair(string text, int openIndex, char openChar, char closeChar, out int closeIndex)
        {
            closeIndex = -1;
            if (string.IsNullOrEmpty(text) || openIndex < 0 || openIndex >= text.Length || text[openIndex] != openChar)
            {
                return false;
            }

            int depth = 0;
            int index = openIndex;
            while (index < text.Length)
            {
                char ch = text[index];
                if (ch == '/' && index + 1 < text.Length)
                {
                    if (text[index + 1] == '/')
                    {
                        index = SkipLineComment(text, index + 2);
                        continue;
                    }

                    if (text[index + 1] == '*')
                    {
                        index = SkipBlockComment(text, index + 2);
                        continue;
                    }
                }

                if (ch == '"' || ch == '\'')
                {
                    index = SkipQuotedLiteral(text, index);
                    continue;
                }

                if (ch == openChar)
                {
                    depth++;
                }
                else if (ch == closeChar)
                {
                    depth--;
                    if (depth == 0)
                    {
                        closeIndex = index;
                        return true;
                    }
                }

                index++;
            }

            return false;
        }

        private static bool TryExtractDirectReturnGetterExpression(string bodyText, out string returnExpression)
        {
            returnExpression = null;
            if (string.IsNullOrWhiteSpace(bodyText))
            {
                return false;
            }

            string compact = StripComments(bodyText).Trim();
            if (string.IsNullOrEmpty(compact) || compact.IndexOf('#') >= 0)
            {
                return false;
            }

            if (!compact.StartsWith("return", StringComparison.Ordinal))
            {
                return false;
            }

            if (compact.Length <= 6 || !char.IsWhiteSpace(compact[6]))
            {
                return false;
            }

            int semicolonIndex = compact.IndexOf(';');
            if (semicolonIndex < 0 || semicolonIndex != compact.Length - 1)
            {
                return false;
            }

            string expression = compact.Substring(6, semicolonIndex - 6).Trim();
            if (string.IsNullOrEmpty(expression))
            {
                return false;
            }

            if (!IsDirectMemberExpression(expression))
            {
                return false;
            }

            returnExpression = expression;
            return true;
        }

        private static bool IsDirectMemberExpression(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return false;
            }

            int index = 0;
            bool expectIdentifier = true;
            while (index < expression.Length)
            {
                index = SkipWhitespace(expression, index);
                if (index >= expression.Length)
                {
                    break;
                }

                if (expectIdentifier)
                {
                    int tokenEnd = ReadWordToken(expression, index);
                    if (tokenEnd <= index)
                    {
                        return false;
                    }

                    index = tokenEnd;
                    expectIdentifier = false;
                    continue;
                }

                if (expression[index] == '.')
                {
                    index++;
                    expectIdentifier = true;
                    continue;
                }

                if (expression[index] == '-' && index + 1 < expression.Length && expression[index + 1] == '>')
                {
                    index += 2;
                    expectIdentifier = true;
                    continue;
                }

                if (expression[index] == ':' && index + 1 < expression.Length && expression[index + 1] == ':')
                {
                    index += 2;
                    expectIdentifier = true;
                    continue;
                }

                return false;
            }

            return !expectIdentifier;
        }

        private static int SkipWhitespace(string text, int startIndex)
        {
            int index = startIndex;
            while (index < text.Length && char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            return index;
        }

        private static int SkipWhitespaceBackward(string text, int startIndex)
        {
            int index = startIndex;
            while (index >= 0 && char.IsWhiteSpace(text[index]))
            {
                index--;
            }

            return index;
        }

        private static int SkipWhitespaceAndComments(string text, int startIndex)
        {
            int index = startIndex;
            while (index >= 0 && index < text.Length)
            {
                if (char.IsWhiteSpace(text[index]))
                {
                    index++;
                    continue;
                }

                if (text[index] == '/' && index + 1 < text.Length)
                {
                    if (text[index + 1] == '/')
                    {
                        index = SkipLineComment(text, index + 2);
                        continue;
                    }

                    if (text[index + 1] == '*')
                    {
                        index = SkipBlockComment(text, index + 2);
                        continue;
                    }
                }

                break;
            }

            return index;
        }

        private static int SkipLineComment(string text, int startIndex)
        {
            int index = startIndex;
            while (index < text.Length && text[index] != '\r' && text[index] != '\n')
            {
                index++;
            }

            return index;
        }

        private static int SkipBlockComment(string text, int startIndex)
        {
            int index = startIndex;
            while (index + 1 < text.Length)
            {
                if (text[index] == '*' && text[index + 1] == '/')
                {
                    return index + 2;
                }

                index++;
            }

            return text.Length;
        }

        private static string StripComments(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(text.Length);
            int index = 0;
            while (index < text.Length)
            {
                if (text[index] == '/' && index + 1 < text.Length)
                {
                    if (text[index + 1] == '/')
                    {
                        index = SkipLineComment(text, index + 2);
                        continue;
                    }

                    if (text[index + 1] == '*')
                    {
                        index = SkipBlockComment(text, index + 2);
                        continue;
                    }
                }

                builder.Append(text[index]);
                index++;
            }

            return builder.ToString();
        }

        private static int ReadWordToken(string text, int startIndex)
        {
            if (startIndex < 0 || startIndex >= text.Length || !(text[startIndex] == '_' || char.IsLetter(text[startIndex])))
            {
                return startIndex;
            }

            int index = startIndex + 1;
            while (index < text.Length && IsWordChar(text[index]))
            {
                index++;
            }

            return index;
        }

        private static string RemoveWhitespace(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var chars = new char[value.Length];
            int count = 0;
            for (int i = 0; i < value.Length; i++)
            {
                char ch = value[i];
                if (!char.IsWhiteSpace(ch))
                {
                    chars[count] = ch;
                    count++;
                }
            }

            return new string(chars, 0, count);
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                return System.IO.Path.GetFullPath(path).TrimEnd('\\', '/');
            }
            catch
            {
                return path.TrimEnd('\\', '/');
            }
        }

        private static bool PathsEqual(string left, string right)
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }

        private static EvaluationPolicy CreateStandardEvaluationPolicy(int configuredPreviousLineCount)
        {
            int clampedPreviousLines = Math.Max(0, configuredPreviousLineCount);
            return new EvaluationPolicy(
                previousLineCount: clampedPreviousLines,
                perExpressionTimeoutMs: DefaultPerExpressionTimeoutMs,
                allowFallbackExpressions: true,
                allowDataMemberRecursion: false,
                allowArrayProbing: false,
                allowCharProbing: false);
        }

        private static EvaluationPolicy CreateManualFunctionSweepPolicy(int configuredPreviousLineCount)
        {
            int clampedPreviousLines = Math.Max(0, configuredPreviousLineCount);
            return new EvaluationPolicy(
                previousLineCount: clampedPreviousLines,
                perExpressionTimeoutMs: 150,
                allowFallbackExpressions: true,
                allowDataMemberRecursion: false,
                allowArrayProbing: false,
                allowCharProbing: false);
        }

        private EvaluationPolicy CreateEvaluationPolicy(int configuredPreviousLineCount, bool manualVariableFunctionSweep)
        {
            if (manualVariableFunctionSweep)
            {
                return CreateManualFunctionSweepPolicy(configuredPreviousLineCount);
            }

            return CreateStandardEvaluationPolicy(configuredPreviousLineCount);
        }

        private static bool ShouldAbortEvaluation()
        {
            return false;
        }

        private static bool RuntimeAllowsFallbackExpressions()
        {
            PerfSession session = currentPerfSession;
            return session == null || session.Policy.AllowFallbackExpressions;
        }

        private static bool RuntimeAllowsDataMemberRecursion()
        {
            PerfSession session = currentPerfSession;
            return session == null || session.Policy.AllowDataMemberRecursion;
        }

        private static bool RuntimeAllowsArrayProbing()
        {
            PerfSession session = currentPerfSession;
            return session == null || session.Policy.AllowArrayProbing;
        }

        private static bool RuntimeAllowsCharProbing()
        {
            PerfSession session = currentPerfSession;
            return session == null || session.Policy.AllowCharProbing;
        }

        private static bool HasEvaluationKind(InlineValueEvaluationKinds evaluationKinds, InlineValueEvaluationKinds requiredKind)
        {
            return (evaluationKinds & requiredKind) == requiredKind;
        }

        private static bool HasRuleKind(InlineValueRuleKinds ruleKinds, InlineValueRuleKinds requiredKind)
        {
            return (ruleKinds & requiredKind) == requiredKind;
        }

        private static CustomRuleDecision GetCustomRuleDecision(
            IReadOnlyList<InlineValueCustomRule> customRules,
            string typeText,
            string expressionText)
        {
            if (customRules == null || customRules.Count == 0)
            {
                return CustomRuleDecision.None;
            }

            string nameText = ExtractTrailingIdentifier(expressionText);
            CustomRuleDecision decision = CustomRuleDecision.None;
            foreach (InlineValueCustomRule rule in customRules)
            {
                if (!rule.Matches(typeText, nameText, expressionText))
                {
                    continue;
                }

                decision = rule.Action == InlineValueCustomRuleAction.Hide
                    ? CustomRuleDecision.Hide
                    : CustomRuleDecision.Show;
            }

            return decision;
        }

        private static EnvDTE.Expression TimedGetExpression(Debugger debugger, string expressionText, int timeout, string phase)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            PerfSession session = currentPerfSession;
            if (session != null)
            {
                timeout = Math.Min(timeout, session.Policy.PerExpressionTimeoutMs);
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            try
            {
                return debugger.GetExpression(expressionText, false, timeout);
            }
            finally
            {
                RecordPerf(phase, expressionText, stopwatch.ElapsedMilliseconds);
            }
        }

        private static void RecordPerf(string phase, string subject, long elapsedMs)
        {
            PerfSession session = currentPerfSession;
            if (session == null)
            {
                return;
            }

            session.Record(phase, subject, elapsedMs);
        }

        private static void WritePerfMessage(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                IVsOutputWindow output = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
                if (output == null)
                {
                    return;
                }

                Guid paneGuid = PerfOutputPaneGuid;
                output.CreatePane(ref paneGuid, "Inline C++ Value Perf", 1, 1);
                output.GetPane(ref paneGuid, out IVsOutputWindowPane pane);
                pane?.OutputStringThreadSafe(message + Environment.NewLine);
            }
            catch
            {
            }
        }

        private static string GetProfileRequestLabel(DebuggerBridge.ProfileRequestKind profileRequestKind)
        {
            switch (profileRequestKind)
            {
                case DebuggerBridge.ProfileRequestKind.GetButton:
                    return "GET button";
                case DebuggerBridge.ProfileRequestKind.ToggleOnButton:
                    return "ON/off button";
                default:
                    return "manual request";
            }
        }

        private Border CreateAdornment(string displayText, int highlightLevel, ValueDisplayKind displayKind)
        {
            Brush background = displayKind == ValueDisplayKind.Uninitialized
                ? uninitializedChipBackground
                : chipBackground;
            if (highlightLevel > 0 && displayKind != ValueDisplayKind.Uninitialized)
            {
                background = new SolidColorBrush(BlendForLevel(highlightLevel));
            }

            Brush foreground;
            if (displayKind == ValueDisplayKind.Null)
            {
                foreground = NullValueForeground;
            }
            else if (displayKind == ValueDisplayKind.Status)
            {
                foreground = StatusForeground;
            }
            else
            {
                foreground = ValueForeground;
            }

            var text = new TextBlock
            {
                Text = " " + displayText,
                Foreground = foreground,
                FontSize = chipFontSize,
                Margin = new Thickness(3, 0, 3, 0),
            };

            var border = new Border
            {
                Background = background,
                BorderBrush = ChipBorder,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(2),
                Margin = new Thickness(2, 0, 0, 0),
                Child = text,
            };

            return border;
        }

        private bool TryResolveManualGetterRequest(Point mousePosition, out string expressionText)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            expressionText = null;
            if (!settings.IsEnabled || !debuggerBridge.TryGetCurrentBreakContext(out DebuggerBridge.BreakContext context))
            {
                return false;
            }

            var line = textView.TextViewLines?.GetTextViewLineContainingYCoordinate(mousePosition.Y);
            if (line == null)
            {
                return false;
            }

            SnapshotPoint? bufferPoint = line.GetBufferPositionFromXCoordinate(mousePosition.X);
            if (!bufferPoint.HasValue || bufferPoint.Value.Snapshot != textBuffer.CurrentSnapshot)
            {
                return false;
            }

            ITextSnapshot snapshot = bufferPoint.Value.Snapshot;
            if (!PathsEqual(normalizedDocumentPath, NormalizePath(context.FileName)))
            {
                return false;
            }

            ITextSnapshotLine snapshotLine = bufferPoint.Value.GetContainingLine();
            if (snapshotLine.LineNumber > context.LineNumber - 1)
            {
                return false;
            }

            string lineText = snapshotLine.GetText();
            int lineOffset = bufferPoint.Value.Position - snapshotLine.Start.Position;
            List<CppCurrentLineTokenizer.IdentifierToken> tokens = CppCurrentLineTokenizer.TokenizeIdentifiers(lineText);
            foreach (CppCurrentLineTokenizer.IdentifierToken token in tokens)
            {
                if (lineOffset < token.Start || lineOffset > token.Start + token.Length)
                {
                    continue;
                }

                if (!TryParseGetterCall(lineText, token, out GetterCallToken getterCall) ||
                    !TryBindVisibleDirectReturnGetter(snapshot, getterCall, out GetterCallToken boundGetterCall))
                {
                    return false;
                }

                expressionText = boundGetterCall.ExpressionText;
                return true;
            }

            return false;
        }

        private void OnBufferChanged(object sender, TextContentChangedEventArgs e)
        {
            Invalidate();
        }

        private void OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (disposed || textView.IsClosed)
            {
                return;
            }

            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
            {
                return;
            }

            ThreadHelper.ThrowIfNotOnUIThread();
            if (!TryResolveManualGetterRequest(e.GetPosition(textView.VisualElement), out string expressionText))
            {
                return;
            }

            manualGetterRequests.Add(expressionText);
            Invalidate();
            e.Handled = true;
        }

        private void OnDebugStateChanged(object sender, EventArgs e)
        {
            manualGetterRequests.Clear();
            Invalidate();
        }

        private void OnSettingsChanged(object sender, EventArgs e)
        {
            UpdateAppearanceSettings();
            Invalidate();
        }

        private void Invalidate()
        {
            cachedSnapshot = null;
            cachedData = null;

            if (disposed || textView.IsClosed)
            {
                return;
            }

            if (textView.VisualElement.Dispatcher.CheckAccess())
            {
                RaiseTagsChanged();
            }
        }

        private void RaiseTagsChanged()
        {
            ITextSnapshot snapshot = textBuffer.CurrentSnapshot;
            if (snapshot.Length == 0)
            {
                TagsChanged?.Invoke(this, new SnapshotSpanEventArgs(new SnapshotSpan(snapshot, 0, 0)));
                return;
            }

            TagsChanged?.Invoke(this, new SnapshotSpanEventArgs(new SnapshotSpan(snapshot, 0, snapshot.Length)));
        }

        private void OnTextViewClosed(object sender, EventArgs e)
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            textBuffer.Changed -= OnBufferChanged;
            textView.Closed -= OnTextViewClosed;
            textView.VisualElement.PreviewMouseLeftButtonDown -= OnPreviewMouseLeftButtonDown;
            debuggerBridge.DebugStateChanged -= OnDebugStateChanged;
            settings.SettingsChanged -= OnSettingsChanged;
        }

        private void UpdateAppearanceSettings()
        {
            chipBackgroundColor = ParseColorOrFallback(settings.ValueBackgroundColor, DefaultChipBackgroundColor);
            uninitializedChipBackgroundColor = ParseColorOrFallback(settings.UninitializedValueBackgroundColor, DefaultUninitializedChipBackgroundColor);
            changedAccentColor = ParseColorOrFallback(settings.ValueChangedAccentColor, DefaultChangedAccentColor);
            chipBackground = CreateFrozenBrush(chipBackgroundColor);
            uninitializedChipBackground = CreateFrozenBrush(uninitializedChipBackgroundColor);
            chipFontSize = settings.ValueChipFontSize;
        }

        private static Color ParseColorOrFallback(string text, Color fallback)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return fallback;
            }

            try
            {
                object parsed = ColorConverter.ConvertFromString(text);
                if (parsed is Color color)
                {
                    return color;
                }
            }
            catch
            {
            }

            return fallback;
        }

        private Dictionary<string, int> GetHighlightLevels(Dictionary<string, string> currentValues, int debuggerVersion)
        {
            if (debuggerVersion != lastComparedDebuggerVersion)
            {
                var next = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (KeyValuePair<string, string> pair in currentValues)
                {
                    if (!lastResolvedValues.TryGetValue(pair.Key, out string previous))
                    {
                        if (lastHighlightLevels.TryGetValue(pair.Key, out int existingLevel) && existingLevel > 1)
                        {
                            next[pair.Key] = existingLevel - 1;
                        }

                        continue;
                    }

                    if (!string.Equals(previous, pair.Value, StringComparison.Ordinal))
                    {
                        next[pair.Key] = MaxManualFadeSteps;
                        continue;
                    }

                    if (lastHighlightLevels.TryGetValue(pair.Key, out int level) && level > 1)
                    {
                        next[pair.Key] = level - 1;
                    }
                }

                lastResolvedValues = new Dictionary<string, string>(currentValues, StringComparer.Ordinal);
                lastHighlightLevels = next;
                lastComparedDebuggerVersion = debuggerVersion;
            }

            return lastHighlightLevels;
        }

        private void ClearChangeTracking()
        {
            lastComparedDebuggerVersion = -1;
            lastResolvedValues.Clear();
            lastHighlightLevels.Clear();
        }

        private Color BlendForLevel(int highlightLevel)
        {
            if (highlightLevel >= MaxManualFadeSteps)
            {
                return changedAccentColor;
            }

            if (highlightLevel <= 0)
            {
                return chipBackgroundColor;
            }

            double t = (MaxManualFadeSteps - highlightLevel) / (double)MaxManualFadeSteps;
            return Blend(changedAccentColor, chipBackgroundColor, t);
        }

        private static Color Blend(Color from, Color to, double t)
        {
            if (t <= 0)
            {
                return from;
            }

            if (t >= 1)
            {
                return to;
            }

            byte a = (byte)(from.A + ((to.A - from.A) * t));
            byte r = (byte)(from.R + ((to.R - from.R) * t));
            byte g = (byte)(from.G + ((to.G - from.G) * t));
            byte b = (byte)(from.B + ((to.B - from.B) * t));
            return Color.FromArgb(a, r, g, b);
        }

        private sealed class PerfSession : IDisposable
        {
            private readonly string filePath;
            private readonly int lineNumber;
            private readonly int debuggerVersion;
            private readonly DebuggerBridge.ProfileRequestKind profileRequestKind;
            private readonly Stopwatch totalStopwatch;
            private readonly Dictionary<string, PerfEntry> entries;

            private PerfSession(
                string filePath,
                int lineNumber,
                int debuggerVersion,
                EvaluationPolicy policy,
                DebuggerBridge.ProfileRequestKind profileRequestKind)
            {
                this.filePath = filePath;
                this.lineNumber = lineNumber;
                this.debuggerVersion = debuggerVersion;
                this.profileRequestKind = profileRequestKind;
                Policy = policy;
                totalStopwatch = Stopwatch.StartNew();
                entries = new Dictionary<string, PerfEntry>(StringComparer.Ordinal);
            }

            public static PerfSession Start(
                string filePath,
                int lineNumber,
                int debuggerVersion,
                EvaluationPolicy policy,
                DebuggerBridge.ProfileRequestKind profileRequestKind)
            {
                return new PerfSession(filePath, lineNumber, debuggerVersion, policy, profileRequestKind);
            }

            public EvaluationPolicy Policy { get; }

            public void Record(string phase, string subject, long elapsedMs)
            {
                if (elapsedMs <= 0)
                {
                    return;
                }

                string key = phase + " | " + NormalizeSubject(subject);
                if (!entries.TryGetValue(key, out PerfEntry entry))
                {
                    entry = new PerfEntry();
                }

                entry.TotalMs += elapsedMs;
                entry.Count++;
                if (elapsedMs > entry.MaxMs)
                {
                    entry.MaxMs = elapsedMs;
                }

                entries[key] = entry;
            }

            public void Dispose()
            {
                totalStopwatch.Stop();
                if (WriteRequestedProfileMessageIfNeeded())
                {
                    return;
                }

                if (totalStopwatch.ElapsedMilliseconds < PerfLogThresholdMs)
                {
                    return;
                }

                var topEntries = entries
                    .Where(pair => pair.Value.MaxMs >= SlowEvaluationThresholdMs || pair.Value.TotalMs >= SlowEvaluationThresholdMs)
                    .OrderByDescending(pair => pair.Value.TotalMs)
                    .ThenByDescending(pair => pair.Value.MaxMs)
                    .Take(8)
                    .ToList();

                if (topEntries.Count == 0)
                {
                    return;
                }

                var builder = new StringBuilder(512);
                builder.Append("InlineCppVarDbg PERF total=");
                builder.Append(totalStopwatch.ElapsedMilliseconds);
                builder.Append("ms");
                builder.Append(", dbgVer=");
                builder.Append(debuggerVersion);
                builder.Append(", file=");
                builder.Append(System.IO.Path.GetFileName(filePath));
                builder.Append(':');
                builder.Append(lineNumber);
                if (topEntries.Count > 0)
                {
                    builder.Append(" | slowest: ");
                }

                for (int i = 0; i < topEntries.Count; i++)
                {
                    KeyValuePair<string, PerfEntry> pair = topEntries[i];
                    if (i > 0)
                    {
                        builder.Append(" ; ");
                    }

                    builder.Append(pair.Key);
                    builder.Append(" total=");
                    builder.Append(pair.Value.TotalMs);
                    builder.Append("ms max=");
                    builder.Append(pair.Value.MaxMs);
                    builder.Append("ms n=");
                    builder.Append(pair.Value.Count);
                }

                string message = builder.ToString();
                System.Diagnostics.Debug.WriteLine(message);
                System.Diagnostics.Trace.WriteLine(message);
                ThreadHelper.ThrowIfNotOnUIThread();
                WritePerfMessage(message);
            }

            private bool WriteRequestedProfileMessageIfNeeded()
            {
                if (profileRequestKind == DebuggerBridge.ProfileRequestKind.None ||
                    totalStopwatch.ElapsedMilliseconds < RequestedProfileLogThresholdMs)
                {
                    return false;
                }

                var orderedEntries = entries
                    .OrderByDescending(pair => pair.Value.TotalMs)
                    .ThenByDescending(pair => pair.Value.MaxMs)
                    .ThenBy(pair => pair.Key, StringComparer.Ordinal)
                    .ToList();

                var builder = new StringBuilder(Math.Max(1024, 96 + (orderedEntries.Count * 80)));
                builder.Append("InlineCppVarDbg PROFILE request=");
                builder.Append(GetProfileRequestLabel(profileRequestKind));
                builder.Append(", total=");
                builder.Append(totalStopwatch.ElapsedMilliseconds);
                builder.Append("ms, dbgVer=");
                builder.Append(debuggerVersion);
                builder.Append(", file=");
                builder.Append(System.IO.Path.GetFileName(filePath));
                builder.Append(':');
                builder.Append(lineNumber);

                if (orderedEntries.Count == 0)
                {
                    builder.AppendLine();
                    builder.Append("  <no profiling entries recorded>");
                }
                else
                {
                    foreach (KeyValuePair<string, PerfEntry> pair in orderedEntries)
                    {
                        builder.AppendLine();
                        builder.Append("  ");
                        builder.Append(pair.Key);
                        builder.Append(" total=");
                        builder.Append(pair.Value.TotalMs);
                        builder.Append("ms max=");
                        builder.Append(pair.Value.MaxMs);
                        builder.Append("ms n=");
                        builder.Append(pair.Value.Count);
                    }
                }

                string message = builder.ToString();
                System.Diagnostics.Debug.WriteLine(message);
                System.Diagnostics.Trace.WriteLine(message);
                ThreadHelper.ThrowIfNotOnUIThread();
                WritePerfMessage(message);
                return true;
            }

            private static string NormalizeSubject(string subject)
            {
                if (string.IsNullOrWhiteSpace(subject))
                {
                    return "<empty>";
                }

                string compact = subject.Replace("\r", " ").Replace("\n", " ").Trim();
                if (compact.Length > 72)
                {
                    return compact.Substring(0, 69) + "...";
                }

                return compact;
            }

            private struct PerfEntry
            {
                public long TotalMs;
                public long MaxMs;
                public int Count;
            }
        }

        private readonly struct EvaluationPolicy
        {
            public EvaluationPolicy(
                int previousLineCount,
                int perExpressionTimeoutMs,
                bool allowFallbackExpressions,
                bool allowDataMemberRecursion,
                bool allowArrayProbing,
                bool allowCharProbing)
            {
                PreviousLineCount = previousLineCount;
                PerExpressionTimeoutMs = perExpressionTimeoutMs;
                AllowFallbackExpressions = allowFallbackExpressions;
                AllowDataMemberRecursion = allowDataMemberRecursion;
                AllowArrayProbing = allowArrayProbing;
                AllowCharProbing = allowCharProbing;
            }

            public int PreviousLineCount { get; }
            public int PerExpressionTimeoutMs { get; }
            public bool AllowFallbackExpressions { get; }
            public bool AllowDataMemberRecursion { get; }
            public bool AllowArrayProbing { get; }
            public bool AllowCharProbing { get; }
        }

        private sealed class ComputedTagData
        {
            public ComputedTagData(SnapshotSpan coveredSpan, List<TagEntry> tags)
            {
                CoveredSpan = coveredSpan;
                Tags = tags;
            }

            public SnapshotSpan CoveredSpan { get; }
            public List<TagEntry> Tags { get; }
        }

        private readonly struct LineTokens
        {
            public LineTokens(
                ITextSnapshotLine line,
                List<CppCurrentLineTokenizer.IdentifierToken> tokens,
                List<GetterCallToken> getterCalls,
                List<GetterCallToken> memberAccesses)
            {
                Line = line;
                Tokens = tokens;
                GetterCalls = getterCalls;
                MemberAccesses = memberAccesses;
            }

            public ITextSnapshotLine Line { get; }
            public List<CppCurrentLineTokenizer.IdentifierToken> Tokens { get; }
            public List<GetterCallToken> GetterCalls { get; }
            public List<GetterCallToken> MemberAccesses { get; }
        }

        private readonly struct InsertionAnchor
        {
            public InsertionAnchor(int position, PositionAffinity affinity)
            {
                Position = position;
                Affinity = affinity;
            }

            public int Position { get; }
            public PositionAffinity Affinity { get; }
        }

        private readonly struct BlockScope
        {
            public BlockScope(int lineNumber, bool isFunctionStart)
            {
                LineNumber = lineNumber;
                IsFunctionStart = isFunctionStart;
            }

            public int LineNumber { get; }
            public bool IsFunctionStart { get; }
        }

        private readonly struct GetterCallToken
        {
            public GetterCallToken(
                string expressionText,
                string evaluationExpressionText,
                string displayLabel,
                string methodName,
                int callEnd,
                InlineValueEvaluationKinds evaluationKind)
            {
                ExpressionText = expressionText;
                EvaluationExpressionText = evaluationExpressionText;
                DisplayLabel = displayLabel;
                MethodName = methodName;
                CallEnd = callEnd;
                EvaluationKind = evaluationKind;
            }

            public string ExpressionText { get; }
            public string EvaluationExpressionText { get; }
            public string DisplayLabel { get; }
            public string MethodName { get; }
            public int CallEnd { get; }
            public InlineValueEvaluationKinds EvaluationKind { get; }

            public GetterCallToken WithEvaluationExpression(string evaluationExpressionText)
            {
                return new GetterCallToken(
                    ExpressionText,
                    evaluationExpressionText,
                    DisplayLabel,
                    MethodName,
                    CallEnd,
                    EvaluationKind);
            }
        }

        private enum ValueDisplayKind
        {
            Normal = 0,
            Null = 1,
            Uninitialized = 2,
            Status = 3,
        }

        private enum EnumValueRenderMode
        {
            SymbolAndInteger = 0,
            IntegerOnly = 1,
        }

        private enum CustomRuleDecision
        {
            None = 0,
            Show = 1,
            Hide = 2,
        }

        private readonly struct TagEntry
        {
            public TagEntry(int position, string displayText, int highlightLevel, PositionAffinity affinity, ValueDisplayKind displayKind)
            {
                Position = position;
                DisplayText = displayText;
                HighlightLevel = highlightLevel;
                Affinity = affinity;
                DisplayKind = displayKind;
            }

            public int Position { get; }
            public string DisplayText { get; }
            public int HighlightLevel { get; }
            public PositionAffinity Affinity { get; }
            public ValueDisplayKind DisplayKind { get; }
        }
    }
}
