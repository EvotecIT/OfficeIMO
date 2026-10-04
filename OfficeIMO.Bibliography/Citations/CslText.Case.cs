namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    internal CslText CapitalizeNoteStart(CultureInfo culture, CancellationToken cancellationToken) {
        bool found = false;
        string html = TransformHtml(Html, (value, protectedCase) => {
            cancellationToken.ThrowIfCancellationRequested();
            if (found) return value;
            for (int index = 0; index < value.Length; index++) {
                if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(value, index);
                if (category != UnicodeCategory.UppercaseLetter && category != UnicodeCategory.LowercaseLetter &&
                    category != UnicodeCategory.TitlecaseLetter && category != UnicodeCategory.ModifierLetter &&
                    category != UnicodeCategory.OtherLetter && category != UnicodeCategory.DecimalDigitNumber &&
                    category != UnicodeCategory.LetterNumber && category != UnicodeCategory.OtherNumber) continue;
                found = true;
                if (protectedCase || category == UnicodeCategory.DecimalDigitNumber || category == UnicodeCategory.LetterNumber ||
                    category == UnicodeCategory.OtherNumber) return value;
                int length = char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1]) ? 2 : 1;
                return value.Substring(0, index) + value.Substring(index, length).ToUpper(culture) + value.Substring(index + length);
            }
            return value;
        }, cancellationToken);
        return new CslText(HtmlPlain(html, cancellationToken), html, Attempted, Rendered);
    }

    private static string ChangeCase(string value, string mode, CultureInfo culture, bool allUpper, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (value.Length == 0) return value;
        if (mode == "lowercase") return value.ToLower(culture);
        if (mode == "uppercase") return value.ToUpper(culture);
        if (mode != "title" && mode != "sentence" && mode != "capitalize-first" && mode != "capitalize-all") return value;

        string source = mode == "sentence" && allUpper ? value.ToLower(culture) : value;
        // Dotted initials remain a single word, so an article such as 'A' inside
        // a name's initials is never mistaken for a title-case stop word.
        CaseWord[] words = CaseWords(source, cancellationToken).ToArray();
        var output = new StringBuilder(source.Length);
        int copied = 0, phraseEnd = -1;
        for (int index = 0; index < words.Length; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            CaseWord word = words[index];
            string wordValue = source.Substring(word.Index, word.Length);
            string lower = wordValue.ToLower(culture);
            string replacement = wordValue;
            if (mode == "title") {
                int phraseLength = CslTitleCaseWords.MatchLength(source, word.Index, lower);
                if (phraseLength > 0) phraseEnd = Math.Max(phraseEnd, word.Index + phraseLength);
                int before = word.Index - 1;
                while (before >= copied && char.IsWhiteSpace(source[before])) before--;
                bool followsColon = before >= 0 && source[before] == ':';
                bool stop = word.Index < phraseEnd && index != 0 && index != words.Length - 1 && !followsColon;
                replacement = stop ? lower : !allUpper && wordValue != lower ? wordValue : Capitalize(lower, culture);
            } else if ((mode == "capitalize-all" || index == 0) && wordValue == lower) {
                replacement = Capitalize(wordValue, culture);
            }
            output.Append(source, copied, word.Index - copied);
            output.Append(replacement);
            copied = word.Index + word.Length;
        }
        return output.Append(source, copied, source.Length - copied).ToString();
    }

    private static string Capitalize(string word, CultureInfo culture) =>
        word.Length == 0 ? word : char.ToUpper(word[0], culture) + word.Substring(1);
}
