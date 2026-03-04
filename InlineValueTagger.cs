using EnvDTE;
using EnvDTE90a;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Tagging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace InlineCppVarDbg
{
    internal sealed class InlineValueTagger : ITagger<IntraTextAdornmentTag>
    {
        private const string DefaultChipBackgroundHex = "#E5E5E5";
        private const string DefaultChangedAccentHex = "#FF0000";
        private const int MaxManualFadeSteps = 5;

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

        private static readonly Color DefaultChipBackgroundColor = ParseColorOrFallback(DefaultChipBackgroundHex, Colors.LightGray);
        private static readonly Color DefaultChangedAccentColor = ParseColorOrFallback(DefaultChangedAccentHex, Colors.Red);
        private static readonly Brush DefaultChipBackground = CreateFrozenBrush(DefaultChipBackgroundColor);
        private static readonly Brush ChipBorder = CreateFrozenBrush("#C8C8C8");
        private static readonly Brush ValueForeground = CreateFrozenBrush("#0B5CAD");

        private readonly IWpfTextView textView;
        private readonly ITextBuffer textBuffer;
        private readonly string documentPath;
        private readonly DebuggerBridge debuggerBridge;
        private readonly InlineValuesSettings settings;
        private readonly string normalizedDocumentPath;
        private Brush chipBackground = DefaultChipBackground;
        private Color chipBackgroundColor = DefaultChipBackgroundColor;
        private Color changedAccentColor = DefaultChangedAccentColor;
        private double chipFontSize = 10.0;

        private ITextSnapshot cachedSnapshot;
        private int cachedDebuggerVersion = -1;
        private ComputedTagData cachedData;
        private bool disposed;
        private int lastComparedDebuggerVersion = -1;
        private Dictionary<string, string> lastResolvedValues = new Dictionary<string, string>(StringComparer.Ordinal);
        private Dictionary<string, int> lastHighlightLevels = new Dictionary<string, int>(StringComparer.Ordinal);

        public InlineValueTagger(IWpfTextView textView, ITextBuffer textBuffer, string documentPath, DebuggerBridge debuggerBridge, InlineValuesSettings settings)
        {
            this.textView = textView;
            this.textBuffer = textBuffer;
            this.documentPath = documentPath;
            this.debuggerBridge = debuggerBridge;
            this.settings = settings;
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
                var adornment = CreateAdornment(entry.DisplayText, entry.HighlightLevel);
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

            int previousLineCount = settings.PreviousLineCount;
            int startLineNumber = Math.Max(0, lineNumber - previousLineCount);
            int endLineNumber = lineNumber;

            var tokensByLine = new List<LineTokens>();
            var allTokens = new List<CppCurrentLineTokenizer.IdentifierToken>();
            var allGetterCalls = new List<GetterCallToken>();
            for (int lineIndex = startLineNumber; lineIndex <= endLineNumber; lineIndex++)
            {
                ITextSnapshotLine line = snapshot.GetLineFromLineNumber(lineIndex);
                string lineText = line.GetText();
                List<CppCurrentLineTokenizer.IdentifierToken> rawTokens = CppCurrentLineTokenizer.TokenizeIdentifiers(lineText);
                var tokens = new List<CppCurrentLineTokenizer.IdentifierToken>(rawTokens.Count);
                var getterCalls = new List<GetterCallToken>();
                foreach (CppCurrentLineTokenizer.IdentifierToken token in rawTokens)
                {
                    if (TryParseGetterCall(lineText, token, out GetterCallToken getterCall))
                    {
                        getterCalls.Add(getterCall);
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

            if (allTokens.Count == 0 && allGetterCalls.Count == 0)
            {
                ITextSnapshotLine startLine = snapshot.GetLineFromLineNumber(startLineNumber);
                ITextSnapshotLine endLine = snapshot.GetLineFromLineNumber(endLineNumber);
                int spanStart = startLine.Start.Position;
                int spanEnd = endLine.End.Position;
                return new ComputedTagData(new SnapshotSpan(snapshot, Span.FromBounds(spanStart, spanEnd)), new List<TagEntry>());
            }

            var valueMap = ResolveValues(context.Debugger, context.StackFrame, allTokens);
            Dictionary<string, string> getterValueMap = ResolveGetterValues(context.Debugger, allGetterCalls);
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
                        displayText = token.Name + ": " + value;
                    }
                    else
                    {
                        insertion = lineTokens.Line.Start.Position + token.Start + token.Length;
                        affinity = PositionAffinity.Successor;
                        displayText = value;
                    }

                    var tagEntry = new TagEntry(insertion, displayText, highlightLevel, affinity);
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
                        displayText = getterCall.DisplayLabel + ": " + getterValue;
                    }
                    else
                    {
                        insertion = lineTokens.Line.Start.Position + getterCall.CallEnd;
                        affinity = PositionAffinity.Successor;
                        displayText = getterValue;
                    }

                    var tagEntry = new TagEntry(insertion, displayText, highlightLevel, affinity);
                    if (!firstEntryByIdentifier.ContainsKey(getterCall.ExpressionText))
                    {
                        firstEntryByIdentifier[getterCall.ExpressionText] = tagEntry;
                    }

                    latestEntryByIdentifier[getterCall.ExpressionText] = tagEntry;
                }
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

            tagEntries = tagEntries
                .OrderBy(t => t.Position)
                .ThenBy(t => t.Affinity == PositionAffinity.Successor ? 1 : 0)
                .ToList();

            ITextSnapshotLine rangeStartLine = snapshot.GetLineFromLineNumber(startLineNumber);
            ITextSnapshotLine rangeEndLine = snapshot.GetLineFromLineNumber(endLineNumber);
            int rangeStart = rangeStartLine.Start.Position;
            int rangeEnd = rangeEndLine.End.Position;

            return new ComputedTagData(new SnapshotSpan(snapshot, Span.FromBounds(rangeStart, rangeEnd)), tagEntries);
        }

        private static Dictionary<string, string> ResolveValues(Debugger debugger, StackFrame2 stackFrame, List<CppCurrentLineTokenizer.IdentifierToken> tokens)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            var requiredIdentifiers = new HashSet<string>(tokens.Select(t => t.Name), StringComparer.Ordinal);
            AddExpressionValues(values, stackFrame.Locals2[true], requiredIdentifiers, 0, debugger);
            AddExpressionValues(values, stackFrame.Arguments2[true], requiredIdentifiers, 0, debugger);

            foreach (string identifier in requiredIdentifiers)
            {
                if (values.ContainsKey(identifier))
                {
                    continue;
                }

                try
                {
                    EnvDTE.Expression expression = debugger.GetExpression(identifier, false, 50);
                    if (expression != null && expression.IsValidValue)
                    {
                        TryAddExpressionValue(values, identifier, expression.Type, expression.Value, requiredIdentifiers, debugger);
                    }
                }
                catch (COMException)
                {
                }
            }

            return values;
        }

        private static Dictionary<string, string> ResolveGetterValues(Debugger debugger, List<GetterCallToken> getterCalls)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (GetterCallToken getterCall in getterCalls)
            {
                if (!seen.Add(getterCall.ExpressionText))
                {
                    continue;
                }

                try
                {
                    EnvDTE.Expression expression = debugger.GetExpression(getterCall.ExpressionText, false, 50);
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
            Debugger debugger)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (expressions == null || depth > 4)
            {
                return;
            }

            foreach (EnvDTE.Expression expression in expressions)
            {
                if (expression == null)
                {
                    continue;
                }

                try
                {
                    if (!expression.IsValidValue)
                    {
                        continue;
                    }

                    TryAddExpressionValue(values, expression.Name, expression.Type, expression.Value, requiredIdentifiers, debugger);
                }
                catch (COMException)
                {
                }

                try
                {
                    AddExpressionValues(values, expression.DataMembers, requiredIdentifiers, depth + 1, debugger);
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
            Debugger debugger)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
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
            if (TryFormatCharPointerOrArrayValue(debugger, evalExpressionName, type, rawValue, normalized, out string charDisplay))
            {
                normalized = charDisplay;
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

            if (cleaned.Length > 80)
            {
                cleaned = cleaned.Substring(0, 77) + "...";
            }

            return cleaned;
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
            displayValue = normalizedValue;
            if (!IsCharPointerOrArrayType(type) && !LooksLikeStringPointerOrArray(type, rawValue, normalizedValue))
            {
                return false;
            }

            if (!TryExtractQuotedStringContent(rawValue, out string content) &&
                !TryExtractQuotedStringContent(normalizedValue, out content) &&
                !TryExtractCharsFromSingleQuotedItems(rawValue, out content) &&
                !TryExtractCharsFromSingleQuotedItems(normalizedValue, out content) &&
                (!ThreadCheckAndTryReadChars(debugger, expressionName, out content)))
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
            if (debugger == null || string.IsNullOrWhiteSpace(expressionName))
            {
                return false;
            }

            foreach (string candidate in BuildIndexExpressionCandidates(expressionName))
            {
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
                string elementExpr = candidateExpression + "[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                EnvDTE.Expression valueExpr;
                try
                {
                    valueExpr = debugger.GetExpression(elementExpr, false, 50);
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
            if (debugger == null || string.IsNullOrWhiteSpace(candidateExpression))
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
                EnvDTE.Expression expression;
                try
                {
                    expression = debugger.GetExpression(probe, false, 50);
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

        private Border CreateAdornment(string displayText, int highlightLevel)
        {
            Brush background = chipBackground;
            if (highlightLevel > 0)
            {
                background = new SolidColorBrush(BlendForLevel(highlightLevel));
            }

            var text = new TextBlock
            {
                Text = " " + displayText,
                Foreground = ValueForeground,
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
            changedAccentColor = ParseColorOrFallback(settings.ValueChangedAccentColor, DefaultChangedAccentColor);
            chipBackground = CreateFrozenBrush(chipBackgroundColor);
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

        private readonly struct TagEntry
        {
            public TagEntry(int position, string displayText, int highlightLevel, PositionAffinity affinity)
            {
                Position = position;
                DisplayText = displayText;
                HighlightLevel = highlightLevel;
                Affinity = affinity;
            }

            public int Position { get; }
            public string DisplayText { get; }
            public int HighlightLevel { get; }
            public PositionAffinity Affinity { get; }
        }
    }
}
