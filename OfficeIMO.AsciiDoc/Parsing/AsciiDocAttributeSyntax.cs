namespace OfficeIMO.AsciiDoc;

/// <summary>Distinguishes quoted attribute values from punctuation in unquoted text.</summary>
internal static class AsciiDocAttributeSyntax {
    internal static bool IsValueQuote(string source, int index, int start, int end,
        System.Threading.CancellationToken cancellationToken = default) {
        char quote = source[index];
        if (quote != '\'' && quote != '"') return false;
        int previous = index - 1;
        while (previous >= start && char.IsWhiteSpace(source[previous])) previous--;
        if (previous >= start && source[previous] != '=' && source[previous] != ',') return false;

        for (int close = index + 1; close < end; close++) {
            if ((close & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (source[close] == '\\') { close++; continue; }
            if (source[close] != quote) continue;
            int after = close + 1;
            while (after < end && char.IsWhiteSpace(source[after])) after++;
            return after == end || source[after] == ',' || source[after] == ']';
        }
        return false;
    }
}
