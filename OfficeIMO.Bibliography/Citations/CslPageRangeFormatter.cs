namespace OfficeIMO.Bibliography;

/// <summary>Formats compatible page endpoints without conflating distinct alphanumeric identifiers.</summary>
internal static class CslPageRangeFormatter {
    internal static string Format(string value, string? format, string delimiter, CancellationToken token, Func<string, string>? renderSource = null, int maximumCharacters = int.MaxValue) {
        if (format == null) return renderSource?.Invoke(value) ?? value;
        token.ThrowIfCancellationRequested();
        if (delimiter.Length == 0) delimiter = "–";
        return RewriteRanges(value, match => FormatRange(value, match, format, delimiter, token), renderSource ?? (source => source), token, maximumCharacters);
    }

    /// <summary>Expands and converts source endpoints, then inserts delimiter text without parsing it.</summary>
    internal static string FormatNumbers(string value, string format, string delimiter, Func<string, string> convert, Func<string, string> renderSource, CancellationToken token, int maximumCharacters) {
        token.ThrowIfCancellationRequested();
        if (delimiter.Length == 0) delimiter = "–";
        return RewriteRanges(value, match => FormatRange(value, match, format, delimiter, token, convert),
            source => RewriteNumbers(source, convert, renderSource, token, maximumCharacters), token, maximumCharacters);
    }

    // Transform source fragments separately, so inserted locale terms and numeral suffixes stay literal.
    private static string RewriteRanges(string value, Func<(int Start, int LeftEnd, int RightStart, int End, int Separator), string> renderToken,
        Func<string, string> renderSource, CancellationToken token, int maximumCharacters) {
        var output = new StringBuilder();
        int position = 0;
        foreach (var match in Ranges(value, token)) {
            token.ThrowIfCancellationRequested();
            Append(renderSource(value.Substring(position, match.Start - position)));
            Append(renderToken(match));
            position = match.End;
        }
        Append(renderSource(value.Substring(position)));
        return output.ToString();

        void Append(string text) => CslNumberSyntax.Append(output, text, 0, text.Length, maximumCharacters);
    }

    private static string RewriteNumbers(string value, Func<string, string> convert, Func<string, string> renderSource,
        CancellationToken token, int maximumCharacters) {
        var output = new StringBuilder();
        int position = 0;
        foreach (var number in CslNumberSyntax.StandaloneNumbers(value, token)) {
            Append(renderSource(value.Substring(position, number.Start - position)));
            Append(convert(value.Substring(number.Start, number.Length)));
            position = number.Start + number.Length;
        }
        Append(renderSource(value.Substring(position)));
        return output.ToString();

        void Append(string text) => CslNumberSyntax.Append(output, text, 0, text.Length, maximumCharacters);
    }

    private static IEnumerable<(int Start, int LeftEnd, int RightStart, int End, int Separator)> Ranges(string value, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        int position = 0;
        while (position < value.Length) {
            CslNumberSyntax.Check(position, token);
            if (!CslNumberSyntax.IsWord(value[position])) { position++; continue; }
            int start = position++;
            while (position < value.Length && CslNumberSyntax.IsWord(value[position])) { CslNumberSyntax.Check(position, token); position++; }
            int leftEnd = position, separator = position;
            CslNumberSyntax.SkipWhitespace(value, ref separator, token);
            if (separator == value.Length || value[separator] != '-' && value[separator] != '–') continue;
            int rightStart = separator + 1;
            CslNumberSyntax.SkipWhitespace(value, ref rightStart, token);
            if (rightStart == value.Length || !CslNumberSyntax.IsWord(value[rightStart])) continue;
            position = rightStart + 1;
            while (position < value.Length && CslNumberSyntax.IsWord(value[position])) { CslNumberSyntax.Check(position, token); position++; }
            yield return (start, leftEnd, rightStart, position, separator);
        }
    }

    private static string FormatRange(string value, (int Start, int LeftEnd, int RightStart, int End, int Separator) match,
        string format, string delimiter, CancellationToken token, Func<string, string>? convert = null) {
        token.ThrowIfCancellationRequested();
        string firstToken = value.Substring(match.Start, match.LeftEnd - match.Start), lastToken = value.Substring(match.RightStart, match.End - match.RightStart);
        if (!TryEndpoint(firstToken, token, out string firstPrefix, out string first, out string firstSuffix) ||
            !TryEndpoint(lastToken, token, out string lastPrefix, out string last, out string lastSuffix)) {
            return CslNumberFormatter.IsRoman(firstToken) && CslNumberFormatter.IsRoman(lastToken) ? firstToken + delimiter + lastToken : value.Substring(match.Start, match.End - match.Start);
        }

        // Different prefixes or suffixes describe distinct identifiers, not an abbreviatable number range.
        if (firstPrefix != lastPrefix || firstSuffix != lastSuffix)
            return RenderEndpoint(firstPrefix, first, firstSuffix, convert) + value[match.Separator] + RenderEndpoint(lastPrefix, last, lastSuffix, convert);
        if (firstSuffix.Length > 0) return firstToken + delimiter + lastToken;

        if (last.Length < first.Length) {
            string expanded = first.Substring(0, first.Length - last.Length) + last;
            // Do not invent a rollover when the supplied abbreviation would precede the start.
            if (CompareDigits(expanded, first, token) < 0)
                return RenderEndpoint(firstPrefix, first, firstSuffix, convert) + delimiter + RenderEndpoint(lastPrefix, last, lastSuffix, convert);
            last = expanded;
        }
        if (format == "expanded" || convert != null && firstPrefix.Length == 0)
            return RenderEndpoint(firstPrefix, first, firstSuffix, convert) + delimiter + RenderEndpoint(lastPrefix, last, lastSuffix, convert);

        int common = 0;
        // A change in digit width makes all endpoint digits significant (999–1001, not 999–001).
        if (first.Length == last.Length && CompareDigits(first, last, token) <= 0) {
            int minimum = format == "minimal" ? 1 : 2;
            bool chicago = format.StartsWith("chicago", StringComparison.Ordinal);
            if (chicago && (first.Length < 3 || IsZero(first[first.Length - 2]) && IsZero(first[first.Length - 1]))) minimum = last.Length;
            else if (chicago && IsZero(first[first.Length - 2])) minimum = 1;
            while (common < last.Length - minimum && first[common] == last[common]) {
                if ((common & 1023) == 0) token.ThrowIfCancellationRequested();
                common++;
            }
            if (chicago && format != "chicago-16" && first.Length == 4 && common == 1) common = 0;
        }
        return firstPrefix + first + delimiter + last.Substring(common);
    }

    private static string RenderEndpoint(string prefix, string number, string suffix, Func<string, string>? convert) {
        if (convert != null && prefix.Length == 0 && suffix.Length == 0) return convert(number);
        return prefix + number + suffix;
    }

    private static bool TryEndpoint(string value, CancellationToken token, out string prefix, out string number, out string suffix) {
        int end = value.Length;
        while (end > 0 && char.IsLetter(value[end - 1])) {
            if ((end & 1023) == 0) token.ThrowIfCancellationRequested();
            end--;
        }
        int start = end;
        while (start > 0 && char.IsDigit(value[start - 1])) {
            if ((start & 1023) == 0) token.ThrowIfCancellationRequested();
            start--;
        }
        prefix = value.Substring(0, start);
        number = value.Substring(start, end - start);
        suffix = value.Substring(end);
        return number.Length > 0 && (prefix.Length == 0 || char.IsLetter(prefix[prefix.Length - 1]));
    }

    private static bool IsZero(char value) => CharUnicodeInfo.GetDecimalDigitValue(value) == 0;

    // Endpoints have equal widths here; compare digit values without narrowing to machine integers
    // or normalizing the source glyphs (for example, fullwidth and Persian decimal digits).
    private static int CompareDigits(string first, string last, CancellationToken token) {
        for (int index = 0; index < first.Length; index++) {
            if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
            int difference = CharUnicodeInfo.GetDecimalDigitValue(first[index]) - CharUnicodeInfo.GetDecimalDigitValue(last[index]);
            if (difference != 0) return difference;
        }
        return 0;
    }
}
