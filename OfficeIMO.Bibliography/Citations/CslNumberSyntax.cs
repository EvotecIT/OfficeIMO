namespace OfficeIMO.Bibliography;

/// <summary>Scans CSL source numerals in linear time with operation cancellation.</summary>
internal static class CslNumberSyntax {
    internal static bool IsWord(char value) => char.IsLetter(value) || char.IsNumber(value);
    internal static bool IsSeparator(char value) => value == '-' || value == '–' || value == ',' || value == '&';
    internal static void Check(int position, CancellationToken token) { if ((position & 1023) == 0) token.ThrowIfCancellationRequested(); }

    internal static bool IsNumeric(string value, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        int position = 0;
        SkipWhitespace(value, ref position, token);
        if (!ReadNumberPart(value, ref position, value.Length, token)) return false;
        while (true) {
            // Separators can occupy every character-offset checkpoint in a long list.
            token.ThrowIfCancellationRequested();
            SkipWhitespace(value, ref position, token);
            if (position == value.Length) return true;
            if (!IsSeparator(value[position++])) return false;
            SkipWhitespace(value, ref position, token);
            if (!ReadNumberPart(value, ref position, value.Length, token)) return false;
        }
    }

    internal static bool IsNumberPart(string value, int start, int end, CancellationToken token) =>
        ReadNumberPart(value, ref start, end, token) && start == end;

    private static bool ReadNumberPart(string value, ref int position, int end, CancellationToken token) {
        while (position < end && char.IsLetter(value[position])) { Check(position, token); position++; }
        int digits = position;
        while (position < end && char.IsDigit(value[position])) { Check(position, token); position++; }
        if (digits == position) return false;
        while (position < end && char.IsLetter(value[position])) { Check(position, token); position++; }
        return true;
    }

    internal static void SkipWhitespace(string value, ref int position, CancellationToken token) {
        while (position < value.Length && char.IsWhiteSpace(value[position])) { Check(position, token); position++; }
    }

    internal static IEnumerable<(int Start, int Length)> StandaloneNumbers(string value, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        int position = 0;
        while (position < value.Length) {
            Check(position, token);
            if (!char.IsDigit(value[position])) { position++; continue; }
            int start = position++;
            while (position < value.Length && char.IsDigit(value[position])) { Check(position, token); position++; }
            if ((start == 0 || !IsWord(value[start - 1])) && (position == value.Length || !IsWord(value[position])))
                yield return (start, position - start);
        }
    }

    internal static string NormalizeSeparators(string value, int maximumCharacters, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var output = new StringBuilder();
        int source = 0, position = 0;
        while (position < value.Length) {
            Check(position, token);
            char separator = value[position];
            if (!IsSeparator(separator)) { position++; continue; }
            int end = position;
            while (end > source && char.IsWhiteSpace(value[end - 1])) { Check(end, token); end--; }
            Append(output, value, source, end - source, maximumCharacters);
            string replacement = separator == ',' ? ", " : separator == '&' ? " & " : separator.ToString();
            Append(output, replacement, 0, replacement.Length, maximumCharacters);
            position++;
            SkipWhitespace(value, ref position, token);
            source = position;
        }
        Append(output, value, source, value.Length - source, maximumCharacters);
        return output.ToString();
    }

    internal static string NormalizeNumericHyphens(string value, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        char[]? output = null;
        for (int position = 1; position + 1 < value.Length; position++) {
            Check(position, token);
            if (value[position] == '-' && char.IsDigit(value[position - 1]) && char.IsDigit(value[position + 1])) {
                if (output == null) output = value.ToCharArray();
                output[position] = '–';
            }
        }
        return output == null ? value : new string(output);
    }

    // Keep existing entity syntax intact; inserted locale terms are never scanned again.
    internal static string LocalizeAmpersands(string value, string andTerm, int maximumCharacters, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var output = new StringBuilder();
        int source = 0;
        for (int position = 0; position < value.Length; position++) {
            Check(position, token);
            if (value[position] != '&') continue;
            int entityEnd = EntityEnd(value, position + 1, token);
            if (entityEnd > 0) {
                // The terminator is skipped by the character loops; poll before advancing past it.
                token.ThrowIfCancellationRequested();
                position = entityEnd; continue;
            }
            Append(output, value, source, position - source, maximumCharacters);
            Append(output, andTerm, 0, andTerm.Length, maximumCharacters);
            source = position + 1;
        }
        if (source == 0) return value;
        Append(output, value, source, value.Length - source, maximumCharacters);
        return output.ToString();
    }

    private static int EntityEnd(string value, int position, CancellationToken token) {
        if (position == value.Length) return 0;
        bool numeric = value[position] == '#';
        bool hex = false;
        if (numeric) {
            position++;
            if (position < value.Length && value[position] == 'x') { position++; hex = true; }
        } else if (!AsciiLetter(value[position])) return 0;
        int start = position;
        while (position < value.Length && (numeric ? hex ? Hex(value[position]) : AsciiDigit(value[position]) :
            AsciiLetter(value[position]) || AsciiDigit(value[position]))) { Check(position, token); position++; }
        return position < value.Length && value[position] == ';' && position - start >= (numeric ? 1 : 2) ? position : 0;
    }

    private static bool AsciiLetter(char value) => value >= 'A' && value <= 'Z' || value >= 'a' && value <= 'z';
    private static bool AsciiDigit(char value) => value >= '0' && value <= '9';
    private static bool Hex(char value) => AsciiDigit(value) || value >= 'A' && value <= 'F' || value >= 'a' && value <= 'f';

    internal static void Append(StringBuilder output, string value, int start, int count, int maximumCharacters) {
        if (count > maximumCharacters - output.Length) throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
        output.Append(value, start, count);
    }
}
