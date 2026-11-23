namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Matches Excel wildcard criteria (<c>*</c>, <c>?</c> and the <c>~</c> escape) against the whole of
        /// <paramref name="text"/>, ignoring case. Runs in place, so callers can test every cell of a range without
        /// building a pattern per comparison.
        /// </summary>
        private static bool MatchesFormulaWildcard(string pattern, string text) =>
            MatchesFormulaWildcard(pattern, text, 0, prefixOnly: false);

        /// <summary>Returns the first index at or after <paramref name="start"/> where the wildcard pattern matches.</summary>
        private static int FindFormulaWildcard(string pattern, string text, int start) {
            for (int index = start; index <= text.Length; index++) {
                if (MatchesFormulaWildcard(pattern, text, index, prefixOnly: true)) return index;
            }
            return -1;
        }

        // Greedy matching with one backtrack point is complete here because '*' matches any run of characters:
        // a later star can always absorb what an earlier backtrack would have tried.
        private static bool MatchesFormulaWildcard(string pattern, string text, int textIndex, bool prefixOnly) {
            int patternIndex = 0;
            int starPattern = -1;
            int starText = -1;
            while (true) {
                if (prefixOnly && patternIndex >= pattern.Length) return true;
                if (textIndex >= text.Length) break;
                if (patternIndex < pattern.Length) {
                    int next = ReadFormulaWildcardToken(pattern, patternIndex, out char kind, out char literal);
                    if (kind == '*') {
                        starPattern = next;
                        starText = textIndex;
                        patternIndex = next;
                        continue;
                    }
                    if (kind == '?' || FormulaWildcardCharactersEqual(literal, text[textIndex])) {
                        patternIndex = next;
                        textIndex++;
                        continue;
                    }
                }
                if (starPattern < 0) return false;
                patternIndex = starPattern;
                textIndex = ++starText;
            }

            while (patternIndex < pattern.Length) {
                patternIndex = ReadFormulaWildcardToken(pattern, patternIndex, out char kind, out _);
                if (kind != '*') return false;
            }
            return true;
        }

        /// <summary>Reads one pattern element: <c>'*'</c>, <c>'?'</c>, or <c>'\0'</c> with a literal character.</summary>
        private static int ReadFormulaWildcardToken(string pattern, int index, out char kind, out char literal) {
            char current = pattern[index];
            if (current == '~' && index + 1 < pattern.Length
                && (pattern[index + 1] == '*' || pattern[index + 1] == '?' || pattern[index + 1] == '~')) {
                kind = '\0';
                literal = pattern[index + 1];
                return index + 2;
            }
            if (current == '*' || current == '?') {
                kind = current;
                literal = '\0';
                return index + 1;
            }
            kind = '\0';
            literal = current;
            return index + 1;
        }

        private static bool FormulaWildcardCharactersEqual(char left, char right) =>
            left == right || char.ToUpperInvariant(left) == char.ToUpperInvariant(right);
    }
}
