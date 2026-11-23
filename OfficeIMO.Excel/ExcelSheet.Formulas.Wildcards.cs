namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Matches Excel wildcard criteria (<c>*</c>, <c>?</c> and the <c>~</c> escape) against the whole of
        /// <paramref name="text"/>, ignoring case. Runs in place, so callers can test every cell of a range without
        /// building a pattern per comparison.
        /// </summary>
        private static bool MatchesFormulaWildcard(string pattern, string text) =>
            MatchFormulaWildcard(pattern, text, 0, prefixOnly: false) == WildcardMatch.Matched;

        /// <summary>Returns the first index at or after <paramref name="start"/> where the wildcard pattern matches.</summary>
        private static int FindFormulaWildcard(string pattern, string text, int start) {
            // Any match starts with the pattern's leading literal run, so candidate starts come from one KMP pass over
            // the text instead of comparing that run again at every position.
            string run = ReadLeadingFormulaWildcardLiterals(pattern);
            if (run.Length == 0) {
                for (int index = start; index <= text.Length; index++) {
                    switch (MatchFormulaWildcard(pattern, text, index, prefixOnly: true)) {
                        case WildcardMatch.Matched: return index;
                        case WildcardMatch.FailedAfterStar: return -1;
                    }
                }
                return -1;
            }

            int[] fallback = BuildFormulaWildcardFallback(run);
            int matched = 0;
            for (int index = start; index < text.Length; index++) {
                char current = char.ToLowerInvariant(text[index]);
                while (matched > 0 && run[matched] != current) matched = fallback[matched - 1];
                if (run[matched] == current) matched++;
                if (matched < run.Length) continue;

                int candidate = index - run.Length + 1;
                switch (MatchFormulaWildcard(pattern, text, candidate, prefixOnly: true)) {
                    case WildcardMatch.Matched: return candidate;
                    case WildcardMatch.FailedAfterStar: return -1;
                }
                matched = fallback[matched - 1];
            }
            return -1;
        }

        /// <summary>Returns the literal characters before the pattern's first <c>*</c> or <c>?</c>, lower-cased.</summary>
        private static string ReadLeadingFormulaWildcardLiterals(string pattern) {
            var run = new System.Text.StringBuilder();
            for (int index = 0; index < pattern.Length;) {
                index = ReadFormulaWildcardToken(pattern, index, out char kind, out char literal);
                if (kind != '\0') break;
                run.Append(char.ToLowerInvariant(literal));
            }
            return run.ToString();
        }

        private static int[] BuildFormulaWildcardFallback(string run) {
            var fallback = new int[run.Length];
            for (int index = 1, length = 0; index < run.Length; index++) {
                while (length > 0 && run[index] != run[length]) length = fallback[length - 1];
                if (run[index] == run[length]) length++;
                fallback[index] = length;
            }
            return fallback;
        }

        private enum WildcardMatch {
            Matched,
            Failed,
            // The literal run before the first '*' matched and the rest found no match anywhere after it. A later start
            // only moves that run further right, so it cannot match either and the search can stop.
            FailedAfterStar
        }

        // Greedy matching with one backtrack point is complete here because '*' matches any run of characters:
        // a later star can always absorb what an earlier backtrack would have tried.
        private static WildcardMatch MatchFormulaWildcard(string pattern, string text, int textIndex, bool prefixOnly) {
            int patternIndex = 0;
            int starPattern = -1;
            int starText = -1;
            while (true) {
                if (prefixOnly && patternIndex >= pattern.Length) return WildcardMatch.Matched;
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
                if (starPattern < 0) return WildcardMatch.Failed;
                patternIndex = starPattern;
                textIndex = ++starText;
            }

            while (patternIndex < pattern.Length) {
                patternIndex = ReadFormulaWildcardToken(pattern, patternIndex, out char kind, out _);
                if (kind != '*') return starPattern < 0 ? WildcardMatch.Failed : WildcardMatch.FailedAfterStar;
            }
            return WildcardMatch.Matched;
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

        // Invariant lower-case equality is the relation the case-insensitive Regex this replaced used: the Kelvin sign
        // matches 'k', while final sigma stays distinct from sigma even though both upper-case to the same letter.
        private static bool FormulaWildcardCharactersEqual(char left, char right) =>
            left == right || char.ToLowerInvariant(left) == char.ToLowerInvariant(right);
    }
}
