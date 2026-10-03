namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    private readonly struct CaseWord {
        internal CaseWord(int index, int length) { Index = index; Length = length; }
        internal int Index { get; }
        internal int Length { get; }
    }

    /// <summary>Preserves dotted initials and apostrophe-connected words in linear time.</summary>
    private static IEnumerable<CaseWord> CaseWords(string source, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        int position = 0, nextCheck = 0;
        while (position < source.Length) {
            CheckCaseScan(position, ref nextCheck, token);
            int start = position;
            if (char.IsLetter(source[position]) && position + 1 < source.Length && source[position + 1] == '.') {
                do {
                    CheckCaseScan(position, ref nextCheck, token);
                    position += 2;
                } while (position + 1 < source.Length && char.IsLetter(source[position]) && source[position + 1] == '.');
            } else if (CaseWordCharacter(source[position])) {
                do {
                    while (position < source.Length && CaseWordCharacter(source[position])) {
                        CheckCaseScan(position, ref nextCheck, token); position++;
                    }
                    if (position + 1 >= source.Length || !CaseApostrophe(source[position]) || !CaseWordCharacter(source[position + 1])) break;
                    position++;
                } while (true);
            } else { position++; continue; }
            yield return new CaseWord(start, position - start);
        }
    }

    private static void CheckCaseScan(int position, ref int nextCheck, CancellationToken token) {
        if (position < nextCheck) return;
        token.ThrowIfCancellationRequested();
        nextCheck = position + 1024;
    }

    private static bool CaseApostrophe(char value) => value == '\'' || value == '’' || value == '`';
    private static bool CaseWordCharacter(char value) {
        UnicodeCategory category = char.GetUnicodeCategory(value);
        return char.IsLetterOrDigit(value) || category == UnicodeCategory.LetterNumber || category == UnicodeCategory.OtherNumber ||
            category == UnicodeCategory.NonSpacingMark || category == UnicodeCategory.SpacingCombiningMark || category == UnicodeCategory.EnclosingMark;
    }
}
