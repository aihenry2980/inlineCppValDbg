using System;
using System.Collections.Generic;
using System.Text;

namespace InlineCppVarDbg
{
    internal enum InlineValueCustomRuleAction
    {
        Show = 0,
        Hide = 1,
    }

    internal enum InlineValueCustomRuleTarget
    {
        Type = 0,
        Name = 1,
        Expression = 2,
    }

    internal sealed class InlineValueCustomRule
    {
        public InlineValueCustomRule(InlineValueCustomRuleAction action, InlineValueCustomRuleTarget target, string pattern)
        {
            Action = action;
            Target = target;
            Pattern = pattern ?? string.Empty;
            normalizedPattern = Pattern.Trim();
        }

        public InlineValueCustomRuleAction Action { get; }
        public InlineValueCustomRuleTarget Target { get; }
        public string Pattern { get; }

        private readonly string normalizedPattern;

        public bool Matches(string typeText, string nameText, string expressionText)
        {
            string candidate;
            switch (Target)
            {
                case InlineValueCustomRuleTarget.Name:
                    candidate = nameText;
                    break;
                case InlineValueCustomRuleTarget.Expression:
                    candidate = expressionText;
                    break;
                default:
                    candidate = typeText;
                    break;
            }

            return WildcardMatches(candidate, normalizedPattern);
        }

        private static bool WildcardMatches(string candidate, string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern))
            {
                return false;
            }

            candidate = candidate ?? string.Empty;
            int candidateIndex = 0;
            int patternIndex = 0;
            int starPatternIndex = -1;
            int starCandidateIndex = -1;

            while (candidateIndex < candidate.Length)
            {
                if (patternIndex < pattern.Length &&
                    (pattern[patternIndex] == '?' ||
                     char.ToUpperInvariant(pattern[patternIndex]) == char.ToUpperInvariant(candidate[candidateIndex])))
                {
                    patternIndex++;
                    candidateIndex++;
                    continue;
                }

                if (patternIndex < pattern.Length && pattern[patternIndex] == '*')
                {
                    starPatternIndex = patternIndex++;
                    starCandidateIndex = candidateIndex;
                    continue;
                }

                if (starPatternIndex >= 0)
                {
                    patternIndex = starPatternIndex + 1;
                    candidateIndex = ++starCandidateIndex;
                    continue;
                }

                return false;
            }

            while (patternIndex < pattern.Length && pattern[patternIndex] == '*')
            {
                patternIndex++;
            }

            return patternIndex == pattern.Length;
        }
    }

    internal static class InlineValueCustomRuleParser
    {
        public static IReadOnlyList<InlineValueCustomRule> Parse(string text)
        {
            var rules = new List<InlineValueCustomRule>();
            if (string.IsNullOrWhiteSpace(text))
            {
                return rules;
            }

            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                if (TryParseRule(line, out InlineValueCustomRule rule))
                {
                    rules.Add(rule);
                }
            }

            return rules;
        }

        public static string Normalize(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                if (line.Length == 0)
                {
                    continue;
                }

                builder.AppendLine(line);
            }

            return builder.ToString().TrimEnd();
        }

        private static bool TryParseRule(string line, out InlineValueCustomRule rule)
        {
            rule = null;

            InlineValueCustomRuleAction action;
            string remainder;
            if (line.StartsWith("+", StringComparison.Ordinal))
            {
                action = InlineValueCustomRuleAction.Show;
                remainder = line.Substring(1).Trim();
            }
            else if (line.StartsWith("-", StringComparison.Ordinal))
            {
                action = InlineValueCustomRuleAction.Hide;
                remainder = line.Substring(1).Trim();
            }
            else
            {
                int firstSpace = line.IndexOf(' ');
                if (firstSpace <= 0)
                {
                    return false;
                }

                string actionText = line.Substring(0, firstSpace).Trim();
                remainder = line.Substring(firstSpace + 1).Trim();
                if (actionText.Equals("show", StringComparison.OrdinalIgnoreCase))
                {
                    action = InlineValueCustomRuleAction.Show;
                }
                else if (actionText.Equals("hide", StringComparison.OrdinalIgnoreCase))
                {
                    action = InlineValueCustomRuleAction.Hide;
                }
                else
                {
                    return false;
                }
            }

            int colonIndex = remainder.IndexOf(':');
            if (colonIndex <= 0 || colonIndex >= remainder.Length - 1)
            {
                return false;
            }

            string targetText = remainder.Substring(0, colonIndex).Trim();
            string pattern = remainder.Substring(colonIndex + 1).Trim();
            if (pattern.Length == 0)
            {
                return false;
            }

            InlineValueCustomRuleTarget target;
            if (targetText.Equals("type", StringComparison.OrdinalIgnoreCase))
            {
                target = InlineValueCustomRuleTarget.Type;
            }
            else if (targetText.Equals("name", StringComparison.OrdinalIgnoreCase))
            {
                target = InlineValueCustomRuleTarget.Name;
            }
            else if (targetText.Equals("expr", StringComparison.OrdinalIgnoreCase) ||
                     targetText.Equals("expression", StringComparison.OrdinalIgnoreCase))
            {
                target = InlineValueCustomRuleTarget.Expression;
            }
            else
            {
                return false;
            }

            rule = new InlineValueCustomRule(action, target, pattern);
            return true;
        }
    }
}
