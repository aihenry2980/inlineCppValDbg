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
        private const int DefaultHardBudgetMs = 100;
        private const int DefaultPerExpressionTimeoutMs = 10;
        private const int EmergencyHardBudgetMs = 60;
        private const int EmergencyPerExpressionTimeoutMs = 5;
        private const int EmergencyModeStepCount = 5;
        private const int FastModeMaxPreviousLines = 20;
        private const int EmergencyModeMaxPreviousLines = 3;
        private const long PerfLogThresholdMs = 250;
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
            "float", "double", "void", "__int8", "__int16", "__int32", "__int64", "size_t", "ptrdiff_t"
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
        private int policyVersion = -1;
        private int emergencyModeStepsRemaining;
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

            AdvanceEvaluationPolicy(debuggerBridge.Version);
            EvaluationPolicy evaluationPolicy = CreateEvaluationPolicy(settings.PreviousLineCount);
            int functionStartLineNumber = FindEnclosingFunctionStartLine(snapshot, lineNumber);
            int previousLineCount = evaluationPolicy.PreviousLineCount;
            int functionScopeStartLine = functionStartLineNumber >= 0 ? functionStartLineNumber : 0;
            int startLineNumber = Math.Max(functionScopeStartLine, lineNumber - previousLineCount);
            int endLineNumber = lineNumber;
            HashSet<string> parameterNames = GetStackFrameParameterNames(context.StackFrame);
            InlineValueNumericDisplayMode numericDisplayMode = settings.NumericDisplayMode;

            var tokensByLine = new List<LineTokens>();
            var allTokens = new List<CppCurrentLineTokenizer.IdentifierToken>();
            var allGetterCalls = new List<GetterCallToken>();
            for (int lineIndex = startLineNumber; lineIndex <= endLineNumber; lineIndex++)
            {
                ITextSnapshotLine line = snapshot.GetLineFromLineNumber(lineIndex);
                if (IsInactivePreprocessorLine(line))
                {
                    continue;
                }

                string lineText = line.GetText();
                List<CppCurrentLineTokenizer.IdentifierToken> rawTokens = CppCurrentLineTokenizer.TokenizeIdentifiers(lineText);
                var tokens = new List<CppCurrentLineTokenizer.IdentifierToken>(rawTokens.Count);
                var getterCalls = new List<GetterCallToken>();
                foreach (CppCurrentLineTokenizer.IdentifierToken token in rawTokens)
                {
                    if (evaluationPolicy.AllowGetterCalls && TryParseGetterCall(lineText, token, out GetterCallToken getterCall))
                    {
                        getterCalls.Add(getterCall);
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

                tokensByLine.Add(new LineTokens(line, tokens, getterCalls));
                allTokens.AddRange(tokens);
                allGetterCalls.AddRange(getterCalls);
            }

            int candidateCount = allTokens.Select(t => t.Name).Distinct(StringComparer.Ordinal).Count() + parameterNames.Count;
            if (evaluationPolicy.AllowGetterCalls)
            {
                candidateCount += allGetterCalls.Select(g => g.ExpressionText).Distinct(StringComparer.Ordinal).Count();
            }

            if (allTokens.Count == 0 && allGetterCalls.Count == 0 && parameterNames.Count == 0)
            {
                ITextSnapshotLine startLine = snapshot.GetLineFromLineNumber(startLineNumber);
                ITextSnapshotLine endLine = snapshot.GetLineFromLineNumber(endLineNumber);
                int spanStart = startLine.Start.Position;
                int spanEnd = endLine.End.Position;
                return new ComputedTagData(new SnapshotSpan(snapshot, Span.FromBounds(spanStart, spanEnd)), new List<TagEntry>());
            }

            Dictionary<string, string> valueMap;
            Dictionary<string, string> getterValueMap;
            bool budgetExceeded;
            using (PerfSession perfSession = PerfSession.Start(documentPath, lineNumber + 1, debuggerBridge.Version, evaluationPolicy))
            {
                currentPerfSession = perfSession;
                try
                {
                    valueMap = ResolveValues(context.Debugger, context.StackFrame, allTokens, parameterNames, numericDisplayMode);
                    getterValueMap = evaluationPolicy.AllowGetterCalls
                        ? ResolveGetterValues(context.Debugger, allGetterCalls, numericDisplayMode)
                        : new Dictionary<string, string>(StringComparer.Ordinal);
                }
                finally
                {
                    currentPerfSession = null;
                }

                budgetExceeded = perfSession.BudgetExceeded;
            }

            if (budgetExceeded)
            {
                emergencyModeStepsRemaining = Math.Max(emergencyModeStepsRemaining, EmergencyModeStepCount);
            }

            foreach (KeyValuePair<string, string> pair in getterValueMap)
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
            }

            int parameterAnchorStartLine = AddMissingParameterAnchorTags(
                snapshot,
                functionStartLineNumber,
                parameterNames,
                valueMap,
                highlightLevels,
                displayMode,
                firstEntryByIdentifier,
                latestEntryByIdentifier);

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

            int coveredStartLineNumber = startLineNumber;
            if (parameterAnchorStartLine != int.MaxValue)
            {
                coveredStartLineNumber = Math.Min(coveredStartLineNumber, parameterAnchorStartLine);
            }

            if (budgetExceeded || evaluationPolicy.IsEmergencyMode)
            {
                string modeLabel = evaluationPolicy.IsEmergencyMode ? "Emergency mode" : "Fast mode";
                string statusText = budgetExceeded
                    ? string.Format(CultureInfo.InvariantCulture, "{0}: {1}/{2} (budget hit)", modeLabel, latestEntryByIdentifier.Count, candidateCount)
                    : string.Format(CultureInfo.InvariantCulture, "{0}: {1}/{2}", modeLabel, latestEntryByIdentifier.Count, candidateCount);
                ITextSnapshotLine currentLine = snapshot.GetLineFromLineNumber(endLineNumber);
                tagEntries.Add(new TagEntry(currentLine.End.Position, statusText, 0, PositionAffinity.Predecessor, ValueDisplayKind.Status));
            }

            tagEntries = tagEntries
                .OrderBy(t => t.Position)
                .ThenBy(t => t.Affinity == PositionAffinity.Successor ? 1 : 0)
                .ToList();

            ITextSnapshotLine rangeStartLine = snapshot.GetLineFromLineNumber(coveredStartLineNumber);
            ITextSnapshotLine rangeEndLine = snapshot.GetLineFromLineNumber(endLineNumber);
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
            InlineValueNumericDisplayMode numericDisplayMode)
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
                AddExpressionValues(values, locals, requiredIdentifiers, 0, debugger, numericDisplayMode, EnumValueRenderMode.SymbolAndInteger);
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
                AddExpressionValues(values, arguments, requiredIdentifiers, 0, debugger, numericDisplayMode, EnumValueRenderMode.SymbolAndInteger);
            }
            catch (COMException)
            {
            }

            if (!RuntimeAllowsFallbackExpressions() || ShouldAbortEvaluation())
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
                        TryAddExpressionValue(values, identifier, expression.Type, expression.Value, requiredIdentifiers, debugger, numericDisplayMode, EnumValueRenderMode.IntegerOnly);
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
            InlineValueNumericDisplayMode numericDisplayMode)
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

                try
                {
                    EnvDTE.Expression expression = TimedGetExpression(debugger, getterCall.ExpressionText, 50, "Getter.GetExpression");
                    if (expression == null || !expression.IsValidValue)
                    {
                        continue;
                    }

                    if (!IsSimpleReturnType(expression.Type))
                    {
                        continue;
                    }

                    string normalized = NormalizeValue(expression.Value);
                    if (string.IsNullOrEmpty(normalized))
                    {
                        continue;
                    }

                    ThreadHelper.ThrowIfNotOnUIThread();
                    if (TryFormatCharPointerOrArrayValue(expression.Type, expression.Value, normalized, out string charDisplay))
                    {
                        normalized = charDisplay;
                    }
                    else if (TryFormatUninitializedValue(expression.Type, expression.Value, normalized, out string uninitDisplay))
                    {
                        normalized = uninitDisplay;
                    }
                    else if (TryFormatNumericValue(numericDisplayMode, expression.Type, normalized, out string numericDisplay))
                    {
                        normalized = numericDisplay;
                    }

                    if (string.IsNullOrEmpty(normalized))
                    {
                        continue;
                    }

                    if (ShouldSuppressValue(expression.Type, normalized))
                    {
                        continue;
                    }

                    values[getterCall.ExpressionText] = normalized;
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
            EnumValueRenderMode enumValueRenderMode)
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
                    TryAddExpressionValue(values, expression.Name, expression.Type, expression.Value, requiredIdentifiers, debugger, numericDisplayMode, enumValueRenderMode);
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
                    AddExpressionValues(values, dataMembers, requiredIdentifiers, depth + 1, debugger, numericDisplayMode, enumValueRenderMode);
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
            EnumValueRenderMode enumValueRenderMode)
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
            if (TryFormatNullPointerValue(type, rawValue, normalized, out string nullDisplay))
            {
                normalized = nullDisplay;
            }
            else if (TryFormatUninitializedValue(type, rawValue, normalized, out string uninitDisplay))
            {
                normalized = uninitDisplay;
            }
            else if (TryFormatEnumValue(debugger, evalExpressionName, type, rawValue, normalized, enumValueRenderMode, numericDisplayMode, out string enumDisplay))
            {
                normalized = enumDisplay;
            }
            else if (TryFormatArrayValue(debugger, evalExpressionName, type, rawValue, normalized, numericDisplayMode, out string arrayDisplay))
            {
                normalized = arrayDisplay;
            }
            else if (TryFormatCharPointerOrArrayValue(debugger, evalExpressionName, type, rawValue, normalized, out string charDisplay))
            {
                normalized = charDisplay;
            }
            else if (TryFormatNumericValue(numericDisplayMode, type, normalized, out string numericDisplay))
            {
                normalized = numericDisplay;
            }

            if (string.IsNullOrEmpty(normalized))
            {
                return;
            }

            if (ShouldSuppressValue(type, normalized))
            {
                return;
            }

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

            bool likelyEnumType = IsEnumType(type);
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
            if (enumValueRenderMode == EnumValueRenderMode.IntegerOnly)
            {
                displayValue = formattedInteger;
                return true;
            }

            string symbolText = ExtractEnumSymbolText(normalizedValue);
            if (string.IsNullOrEmpty(symbolText))
            {
                symbolText = ExtractEnumSymbolText(rawValue);
            }

            if (string.IsNullOrEmpty(symbolText))
            {
                displayValue = formattedInteger;
                return true;
            }

            displayValue = symbolText + " (" + formattedInteger + ")";
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

        private static bool IsEnumType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            if (ContainsTypeWord(lower, "enum"))
            {
                return true;
            }

            if (type.IndexOf('*') >= 0 || type.IndexOf('&') >= 0 || type.IndexOf('[') >= 0 || type.IndexOf(']') >= 0)
            {
                return false;
            }

            if (ContainsTypeWord(lower, "class") || ContainsTypeWord(lower, "struct") || ContainsTypeWord(lower, "union"))
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
            bool hasCandidateToken = false;
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

                hasCandidateToken = true;
            }

            return hasCandidateToken;
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

        private static bool IsIntegralScalarType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            if (type.IndexOf('*') >= 0 || type.IndexOf('&') >= 0 || type.IndexOf('[') >= 0 || type.IndexOf(']') >= 0)
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

            return ContainsTypeWord(lower, "char") ||
                   ContainsTypeWord(lower, "short") ||
                   ContainsTypeWord(lower, "int") ||
                   ContainsTypeWord(lower, "long") ||
                   ContainsTypeWord(lower, "__int8") ||
                   ContainsTypeWord(lower, "__int16") ||
                   ContainsTypeWord(lower, "__int32") ||
                   ContainsTypeWord(lower, "__int64") ||
                   ContainsTypeWord(lower, "size_t") ||
                   ContainsTypeWord(lower, "ptrdiff_t");
        }

        private static bool ShouldSuppressValue(string type, string normalizedValue)
        {
            if (IsFunctionPointerType(type))
            {
                return true;
            }

            if (IsStructValueType(type, normalizedValue))
            {
                return true;
            }

            if (!IsAddressLikeValue(normalizedValue))
            {
                return false;
            }

            return IsLikelyClassOrStructPointerType(type);
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

        private static bool TryFormatNullPointerValue(string type, string rawValue, string normalizedValue, out string displayValue)
        {
            displayValue = normalizedValue;
            if (!IsPointerType(type))
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

            if (!IsPointerType(type) && !IsIntegralScalarType(type))
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

            // Array value exists but we could not resolve concrete entries without showing an address.
            displayValue = "...";
            return true;
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

        private static bool IsSimpleReturnType(string type)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                return false;
            }

            if (type.IndexOf('*') >= 0 || type.IndexOf('&') >= 0 || type.IndexOf('[') >= 0 || type.IndexOf(']') >= 0)
            {
                return false;
            }

            string lower = type.ToLowerInvariant();
            if (ContainsTypeWord(lower, "class") || ContainsTypeWord(lower, "struct") || ContainsTypeWord(lower, "union"))
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
            bool sawPrimitive = false;
            foreach (string token in tokens)
            {
                if (TypeNoiseTokens.Contains(token))
                {
                    continue;
                }

                if (string.Equals(token, "void", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (PrimitiveTypeTokens.Contains(token))
                {
                    sawPrimitive = true;
                    continue;
                }

                return false;
            }

            return sawPrimitive;
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
            if (!token.Name.StartsWith("get", StringComparison.OrdinalIgnoreCase))
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

            getterCall = new GetterCallToken(expressionText, expressionText, closeParenIndex + 1);
            return true;
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

                if (leftEnd < 0 || !IsWordChar(lineText[leftEnd]))
                {
                    break;
                }

                int leftStart = leftEnd;
                while (leftStart >= 0 && IsWordChar(lineText[leftStart]))
                {
                    leftStart--;
                }

                start = leftStart + 1;
            }

            return start;
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
            if (!memberName.StartsWith("get", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            while (i < lineText.Length && char.IsWhiteSpace(lineText[i]))
            {
                i++;
            }

            return i < lineText.Length && lineText[i] == '(';
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

        private void AdvanceEvaluationPolicy(int debuggerVersion)
        {
            if (policyVersion == debuggerVersion)
            {
                return;
            }

            policyVersion = debuggerVersion;
            if (emergencyModeStepsRemaining > 0)
            {
                emergencyModeStepsRemaining--;
            }
        }

        private static EvaluationPolicy CreateEvaluationPolicy(int configuredPreviousLineCount, bool emergencyMode = false)
        {
            int clampedPreviousLines = Math.Max(0, configuredPreviousLineCount);
            if (!emergencyMode)
            {
                return new EvaluationPolicy(
                    isEmergencyMode: false,
                    previousLineCount: Math.Min(clampedPreviousLines, FastModeMaxPreviousLines),
                    hardBudgetMs: DefaultHardBudgetMs,
                    perExpressionTimeoutMs: DefaultPerExpressionTimeoutMs,
                    allowGetterCalls: true,
                    allowFallbackExpressions: false,
                    allowDataMemberRecursion: false,
                    allowArrayProbing: false,
                    allowCharProbing: false);
            }

            return new EvaluationPolicy(
                isEmergencyMode: true,
                previousLineCount: Math.Min(clampedPreviousLines, EmergencyModeMaxPreviousLines),
                hardBudgetMs: EmergencyHardBudgetMs,
                perExpressionTimeoutMs: EmergencyPerExpressionTimeoutMs,
                allowGetterCalls: true,
                allowFallbackExpressions: false,
                allowDataMemberRecursion: false,
                allowArrayProbing: false,
                allowCharProbing: false);
        }

        private EvaluationPolicy CreateEvaluationPolicy(int configuredPreviousLineCount)
        {
            return CreateEvaluationPolicy(configuredPreviousLineCount, emergencyModeStepsRemaining > 0);
        }

        private static bool ShouldAbortEvaluation()
        {
            PerfSession session = currentPerfSession;
            return session != null && session.ShouldAbort();
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

        private static EnvDTE.Expression TimedGetExpression(Debugger debugger, string expressionText, int timeout, string phase)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            PerfSession session = currentPerfSession;
            if (session != null)
            {
                if (session.ShouldAbort())
                {
                    return null;
                }

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

        private void OnBufferChanged(object sender, TextContentChangedEventArgs e)
        {
            Invalidate();
        }

        private void OnDebugStateChanged(object sender, EventArgs e)
        {
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
            private readonly Stopwatch totalStopwatch;
            private readonly Dictionary<string, PerfEntry> entries;
            private bool budgetExceeded;

            private PerfSession(string filePath, int lineNumber, int debuggerVersion, EvaluationPolicy policy)
            {
                this.filePath = filePath;
                this.lineNumber = lineNumber;
                this.debuggerVersion = debuggerVersion;
                Policy = policy;
                totalStopwatch = Stopwatch.StartNew();
                entries = new Dictionary<string, PerfEntry>(StringComparer.Ordinal);
            }

            public static PerfSession Start(string filePath, int lineNumber, int debuggerVersion, EvaluationPolicy policy)
            {
                return new PerfSession(filePath, lineNumber, debuggerVersion, policy);
            }

            public EvaluationPolicy Policy { get; }

            public bool BudgetExceeded
            {
                get { return budgetExceeded; }
            }

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

            public bool ShouldAbort()
            {
                if (budgetExceeded)
                {
                    return true;
                }

                if (Policy.HardBudgetMs <= 0)
                {
                    return false;
                }

                if (totalStopwatch.ElapsedMilliseconds < Policy.HardBudgetMs)
                {
                    return false;
                }

                budgetExceeded = true;
                return true;
            }

            public void Dispose()
            {
                totalStopwatch.Stop();
                if (totalStopwatch.ElapsedMilliseconds < PerfLogThresholdMs && !budgetExceeded)
                {
                    return;
                }

                var topEntries = entries
                    .Where(pair => pair.Value.MaxMs >= SlowEvaluationThresholdMs || pair.Value.TotalMs >= SlowEvaluationThresholdMs)
                    .OrderByDescending(pair => pair.Value.TotalMs)
                    .ThenByDescending(pair => pair.Value.MaxMs)
                    .Take(8)
                    .ToList();

                if (topEntries.Count == 0 && !budgetExceeded)
                {
                    return;
                }

                var builder = new StringBuilder(512);
                builder.Append("InlineCppVarDbg PERF total=");
                builder.Append(totalStopwatch.ElapsedMilliseconds);
                builder.Append("ms / budget=");
                builder.Append(Policy.HardBudgetMs);
                builder.Append("ms");
                if (budgetExceeded)
                {
                    builder.Append(" [BUDGET EXCEEDED]");
                }
                builder.Append(", mode=");
                builder.Append(Policy.IsEmergencyMode ? "Emergency" : "Fast");
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
                bool isEmergencyMode,
                int previousLineCount,
                int hardBudgetMs,
                int perExpressionTimeoutMs,
                bool allowGetterCalls,
                bool allowFallbackExpressions,
                bool allowDataMemberRecursion,
                bool allowArrayProbing,
                bool allowCharProbing)
            {
                IsEmergencyMode = isEmergencyMode;
                PreviousLineCount = previousLineCount;
                HardBudgetMs = hardBudgetMs;
                PerExpressionTimeoutMs = perExpressionTimeoutMs;
                AllowGetterCalls = allowGetterCalls;
                AllowFallbackExpressions = allowFallbackExpressions;
                AllowDataMemberRecursion = allowDataMemberRecursion;
                AllowArrayProbing = allowArrayProbing;
                AllowCharProbing = allowCharProbing;
            }

            public bool IsEmergencyMode { get; }
            public int PreviousLineCount { get; }
            public int HardBudgetMs { get; }
            public int PerExpressionTimeoutMs { get; }
            public bool AllowGetterCalls { get; }
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
                List<GetterCallToken> getterCalls)
            {
                Line = line;
                Tokens = tokens;
                GetterCalls = getterCalls;
            }

            public ITextSnapshotLine Line { get; }
            public List<CppCurrentLineTokenizer.IdentifierToken> Tokens { get; }
            public List<GetterCallToken> GetterCalls { get; }
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
            public GetterCallToken(string expressionText, string displayLabel, int callEnd)
            {
                ExpressionText = expressionText;
                DisplayLabel = displayLabel;
                CallEnd = callEnd;
            }

            public string ExpressionText { get; }
            public string DisplayLabel { get; }
            public int CallEnd { get; }
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
