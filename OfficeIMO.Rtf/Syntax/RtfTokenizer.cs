using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Tokenizes RTF text into braces, control words, control symbols, text, and binary payloads.
/// </summary>
public static class RtfTokenizer {
    /// <summary>
    /// Tokenizes RTF content.
    /// </summary>
    public static RtfTokenizeResult Tokenize(string rtf) {
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));

        var tokens = new List<RtfToken>();
        var diagnostics = new List<RtfDiagnostic>();
        int position = 0;

        while (position < rtf.Length) {
            char current = rtf[position];
            if (current == '{') {
                tokens.Add(new RtfToken(RtfTokenKind.GroupStart, position, "{"));
                position++;
                continue;
            }

            if (current == '}') {
                tokens.Add(new RtfToken(RtfTokenKind.GroupEnd, position, "}"));
                position++;
                continue;
            }

            if (current == '\\') {
                ReadControl(rtf, tokens, diagnostics, ref position);
                continue;
            }

            int start = position;
            while (position < rtf.Length && rtf[position] != '{' && rtf[position] != '}' && rtf[position] != '\\') {
                position++;
            }

            tokens.Add(new RtfToken(RtfTokenKind.Text, start, rtf.Substring(start, position - start), text: rtf.Substring(start, position - start)));
        }

        tokens.Add(new RtfToken(RtfTokenKind.EndOfFile, position, string.Empty));
        return new RtfTokenizeResult(tokens.AsReadOnly(), diagnostics.AsReadOnly());
    }

    private static void ReadControl(string rtf, List<RtfToken> tokens, List<RtfDiagnostic> diagnostics, ref int position) {
        int start = position;
        position++;
        if (position >= rtf.Length) {
            diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF001", "A trailing backslash does not form a valid control word or control symbol.", start));
            tokens.Add(new RtfToken(RtfTokenKind.ControlSymbol, start, "\\", controlSymbol: '\\'));
            return;
        }

        char current = rtf[position];
        if (IsAsciiLetter(current)) {
            ReadControlWord(rtf, tokens, diagnostics, start, ref position);
            return;
        }

        if (current == '\'' && position + 2 < rtf.Length && IsHexDigit(rtf[position + 1]) && IsHexDigit(rtf[position + 2])) {
            string raw = rtf.Substring(start, 4);
            int value = int.Parse(rtf.Substring(position + 1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            tokens.Add(new RtfToken(RtfTokenKind.ControlSymbol, start, raw, controlSymbol: '\'', parameter: value, hasParameter: true));
            position += 3;
            return;
        }

        tokens.Add(new RtfToken(RtfTokenKind.ControlSymbol, start, rtf.Substring(start, 2), controlSymbol: current));
        position++;
    }

    private static void ReadControlWord(string rtf, List<RtfToken> tokens, List<RtfDiagnostic> diagnostics, int start, ref int position) {
        int nameStart = position;
        while (position < rtf.Length && IsAsciiLetter(rtf[position])) {
            position++;
        }

        string name = rtf.Substring(nameStart, position - nameStart);
        bool negative = false;
        bool hasParameter = false;
        int parameter = 0;

        if (position < rtf.Length && rtf[position] == '-') {
            negative = true;
            position++;
        }

        int digitStart = position;
        while (position < rtf.Length && char.IsDigit(rtf[position])) {
            position++;
        }

        if (position > digitStart) {
            hasParameter = true;
            string digits = rtf.Substring(digitStart, position - digitStart);
            if (!int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out parameter)) {
                diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF002", $"Control word '{name}' has a parameter outside Int32 range.", start));
                parameter = 0;
            }

            if (negative) {
                parameter = -parameter;
            }
        }

        if (position < rtf.Length && rtf[position] == ' ') {
            position++;
        }

        string raw = rtf.Substring(start, position - start);
        tokens.Add(new RtfToken(RtfTokenKind.ControlWord, start, raw, controlName: name, parameter: hasParameter ? parameter : null, hasParameter: hasParameter));

        if (string.Equals(name, "bin", StringComparison.Ordinal) && hasParameter) {
            ReadBinaryPayload(rtf, tokens, diagnostics, start, parameter, ref position);
        }
    }

    private static void ReadBinaryPayload(string rtf, List<RtfToken> tokens, List<RtfDiagnostic> diagnostics, int controlPosition, int length, ref int position) {
        if (length < 0) {
            diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF003", "The \\bin control word cannot declare a negative byte count.", controlPosition));
            return;
        }

        int available = Math.Min(length, rtf.Length - position);
        if (available < length) {
            diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF004", "The \\bin payload ended before the declared byte count was satisfied.", position));
        }

        var data = new byte[available];
        for (int i = 0; i < available; i++) {
            data[i] = (byte)(rtf[position + i] & 0xFF);
        }

        tokens.Add(new RtfToken(RtfTokenKind.Binary, position, rtf.Substring(position, available), binaryData: data));
        position += available;
    }

    private static bool IsAsciiLetter(char value) => (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z');

    private static bool IsHexDigit(char value) =>
        (value >= '0' && value <= '9') ||
        (value >= 'a' && value <= 'f') ||
        (value >= 'A' && value <= 'F');
}
