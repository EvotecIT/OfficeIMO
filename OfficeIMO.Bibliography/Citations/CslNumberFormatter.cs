namespace OfficeIMO.Bibliography;

/// <summary>Transforms source numerals and separators without reparsing generated locale terms.</summary>
internal static class CslNumberFormatter {
    internal static string Format(string value, Func<string, string> convert, string andTerm, bool locator, bool numericHyphens,
        int maximumCharacters, CancellationToken token) {
        var output = new StringBuilder();
        token.ThrowIfCancellationRequested();
        int position = 0, source = 0;
        while (position < value.Length) {
            CslNumberSyntax.Check(position, token);
            int start = position;
            char current = value[position++];
            string rendered;
            if (char.IsDigit(current)) {
                while (position < value.Length && char.IsDigit(value[position])) { CslNumberSyntax.Check(position, token); position++; }
                if (start > 0 && CslNumberSyntax.IsWord(value[start - 1]) || position < value.Length && CslNumberSyntax.IsWord(value[position])) continue;
                rendered = convert(value.Substring(start, position - start));
            } else if (CslNumberSyntax.IsSeparator(current)) {
                rendered = current == '&' ? andTerm : current == '-' && (locator || numericHyphens && start > 0 && position < value.Length &&
                    char.IsDigit(value[start - 1]) && char.IsDigit(value[position])) ? "–" : current.ToString();
            } else continue;
            Append(value, source, start - source);
            Append(rendered, 0, rendered.Length);
            source = position;
        }
        Append(value, source, value.Length - source);
        return output.ToString();

        void Append(string source, int start, int count) {
            if (count > maximumCharacters - output.Length)
                throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
            output.Append(source, start, count);
        }
    }

    internal static bool IsPlural(string value, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        int digitRuns = 0;
        bool previousDigit = false;
        for (int position = 0; position < value.Length; position++) {
            CslNumberSyntax.Check(position, token);
            bool digit = char.IsDigit(value[position]);
            // A decimal or dotted version denotes one value; separators start another value.
            bool dottedContinuation = position > 1 && value[position - 1] == '.' && char.IsDigit(value[position - 2]);
            if (digit && !previousDigit && !dottedContinuation && ++digitRuns > 1) return true;
            previousDigit = digit;
        }
        int source = 0, parts = 0;
        for (int position = 0; position <= value.Length; position++) {
            CslNumberSyntax.Check(position, token);
            if (position < value.Length && !CslNumberSyntax.IsSeparator(value[position])) continue;
            int start = source, end = position;
            while (start < end && char.IsWhiteSpace(value[start])) { CslNumberSyntax.Check(start, token); start++; }
            while (end > start && char.IsWhiteSpace(value[end - 1])) { CslNumberSyntax.Check(end, token); end--; }
            if (!IsRoman(value, start, end) && !CslNumberSyntax.IsNumberPart(value, start, end, token)) return false;
            parts++; source = position + 1;
        }
        return parts > 1;
    }

    internal static bool IsRoman(string value) => IsRoman(value, 0, value.Length);

    private static bool IsRoman(string value, int start, int end) {
        // Canonical Roman numerals in the supported 1..3999 range have at most 15 ASCII letters.
        if (end == start || end - start > 15) return false;
        int position = start;
        for (int count = 0; count < 3 && position < end && char.ToUpperInvariant(value[position]) == 'M'; count++) position++;
        Group('C', 'D', 'M'); Group('X', 'L', 'C'); Group('I', 'V', 'X');
        return position == end;

        void Group(char one, char five, char ten) {
            if (position < end && char.ToUpperInvariant(value[position]) == one && position + 1 < end &&
                (char.ToUpperInvariant(value[position + 1]) == five || char.ToUpperInvariant(value[position + 1]) == ten)) { position += 2; return; }
            if (position < end && char.ToUpperInvariant(value[position]) == five) position++;
            for (int count = 0; count < 3 && position < end && char.ToUpperInvariant(value[position]) == one; count++) position++;
        }
    }
}
