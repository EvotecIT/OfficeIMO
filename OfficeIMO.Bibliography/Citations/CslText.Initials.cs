namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    /// <summary>Initializes visible name text while preserving the markup around retained letters.</summary>
    internal CslText InitializeName(string suffix, bool hyphen, bool initialize, int maximumCharacters, CancellationToken token) {
        string plain = Plain;
        var replacements = new List<(int Start, int Length, string Value, bool AtEnd)>();
        int pending = -1;
        bool pendingHyphen = false, pendingCompound = false, previousInitial = false, hasText = false;
        string initialSuffix = suffix.TrimEnd();
        string initialSpacing = suffix.Substring(initialSuffix.Length);
        long outputLength = 0;
        for (int offset = 0; offset < plain.Length;) {
            token.ThrowIfCancellationRequested();
            int start = offset;
            int kind = NameTextKind(plain[offset++]);
            while (offset < plain.Length && NameTextKind(plain[offset]) == kind) {
                if ((offset & 1023) == 0) token.ThrowIfCancellationRequested();
                offset++;
            }
            int length = offset - start;
            string part = plain.Substring(start, length);
            if (part.All(character => char.IsWhiteSpace(character) || character == '-' || character == '.')) {
                if (pending < 0) pending = replacements.Count;
                pendingHyphen |= hyphen && part.IndexOf('-') >= 0;
                pendingCompound |= part.IndexOf('-') >= 0;
                replacements.Add((start, length, string.Empty, true));
                continue;
            }
            string replacement = part;
            if (char.IsLetter(part, 0)) {
                bool existingInitial = part.Length == (char.IsSurrogatePair(part, 0) ? 2 : 1) ||
                    offset < plain.Length && plain[offset] == '.';
                bool currentInitial = existingInitial || initialize && (char.IsUpper(part, 0) || pendingCompound && previousInitial);
                string separator = !hasText ? string.Empty : pendingHyphen ? "-" :
                    previousInitial && currentInitial ? initialSpacing : " ";
                if (pending >= 0) {
                    var delimiter = replacements[pending];
                    replacements[pending] = (delimiter.Start, delimiter.Length, separator, true);
                } else replacement = separator + replacement;
                if (currentInitial) {
                    string first = existingInitial ? part : part.Substring(0, char.IsSurrogatePair(part, 0) ? 2 : 1);
                    replacement = (pending < 0 ? separator : string.Empty) + first + initialSuffix;
                }
                if (pending >= 0) outputLength += separator.Length;
                previousInitial = currentInitial;
            }
            outputLength += replacement.Length;
            if (outputLength > maximumCharacters) throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
            replacements.Add((start, length, replacement, false));
            hasText = true;
            pending = -1; pendingHyphen = false; pendingCompound = false;
        }

        int sourceOffset = 0, rangeIndex = 0;
        long encodedLength = 0;
        string html = TransformHtml(Html, (text, _) => {
            var result = new StringBuilder(text.Length);
            for (int index = 0; index < text.Length; index++, sourceOffset++) {
                if ((index & 4095) == 0) token.ThrowIfCancellationRequested();
                while (rangeIndex < replacements.Count && sourceOffset >= replacements[rangeIndex].Start + replacements[rangeIndex].Length) rangeIndex++;
                if (rangeIndex == replacements.Count || sourceOffset < replacements[rangeIndex].Start) result.Append(text[index]);
                else if (sourceOffset == replacements[rangeIndex].Start + (replacements[rangeIndex].AtEnd ? replacements[rangeIndex].Length - 1 : 0))
                    result.Append(replacements[rangeIndex].Value);
            }
            string value = result.ToString();
            encodedLength += Escape(value).Length;
            if (encodedLength > maximumCharacters) throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
            return value;
        }, token);
        if (html.Length > maximumCharacters) throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
        return new CslText(HtmlPlain(html, token), html, Attempted, Rendered);
    }

    // Match the original letter/mark, delimiter, and other UTF-16 token classes
    // without a wall-clock regex timeout. Scanning checks cancellation inside a
    // single long token as well as between tokens.
    private static int NameTextKind(char character) {
        UnicodeCategory category = char.GetUnicodeCategory(character);
        if (char.IsLetter(character) || category == UnicodeCategory.NonSpacingMark ||
            category == UnicodeCategory.SpacingCombiningMark || category == UnicodeCategory.EnclosingMark) return 0;
        return character == '.' || character == '-' || char.IsWhiteSpace(character) ? 1 : 2;
    }
}
