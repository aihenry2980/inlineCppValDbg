using System;
using System.Collections.Generic;

namespace InlineCppVarDbg
{
    internal static class CppCurrentLineTokenizer
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "alignas", "alignof", "and", "and_eq", "asm", "auto", "bitand", "bitor",
            "bool", "break", "case", "catch", "char", "char8_t", "char16_t", "char32_t",
            "class", "compl", "concept", "const", "consteval", "constexpr", "constinit",
            "const_cast", "continue", "co_await", "co_return", "co_yield", "decltype",
            "default", "delete", "do", "double", "dynamic_cast", "else", "enum", "explicit",
            "export", "extern", "false", "float", "for", "friend", "goto", "if", "inline",
            "int", "long", "mutable", "namespace", "new", "noexcept", "not", "not_eq",
            "nullptr", "operator", "or", "or_eq", "private", "protected", "public",
            "register", "reinterpret_cast", "requires", "return", "short", "signed",
            "sizeof", "static", "static_assert", "static_cast", "struct", "switch",
            "template", "this", "thread_local", "throw", "true", "try", "typedef",
            "typeid", "typename", "union", "unsigned", "using", "virtual", "void",
            "volatile", "wchar_t", "while", "xor", "xor_eq"
        };

        public static List<IdentifierToken> TokenizeIdentifiers(string line)
        {
            var results = new List<IdentifierToken>();
            if (string.IsNullOrEmpty(line))
            {
                return results;
            }

            int length = line.Length;
            int i = 0;
            while (i < length)
            {
                char ch = line[i];

                if (ch == '/' && i + 1 < length)
                {
                    char next = line[i + 1];
                    if (next == '/')
                    {
                        break;
                    }

                    if (next == '*')
                    {
                        break;
                    }
                }

                if (ch == '"' || ch == '\'')
                {
                    i = SkipQuotedLiteral(line, i);
                    continue;
                }

                if (IsIdentifierStart(ch))
                {
                    int start = i;
                    i++;
                    while (i < length && IsIdentifierPart(line[i]))
                    {
                        i++;
                    }

                    string identifier = line.Substring(start, i - start);
                    if (!Keywords.Contains(identifier))
                    {
                        results.Add(new IdentifierToken(identifier, start, i - start));
                    }

                    continue;
                }

                i++;
            }

            return results;
        }

        private static bool IsIdentifierStart(char ch)
        {
            return ch == '_' || char.IsLetter(ch);
        }

        private static bool IsIdentifierPart(char ch)
        {
            return ch == '_' || char.IsLetterOrDigit(ch);
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

        internal readonly struct IdentifierToken
        {
            public IdentifierToken(string name, int start, int length)
            {
                Name = name;
                Start = start;
                Length = length;
            }

            public string Name { get; }
            public int Start { get; }
            public int Length { get; }
        }
    }
}
