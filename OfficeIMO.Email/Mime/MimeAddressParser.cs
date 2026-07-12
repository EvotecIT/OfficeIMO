namespace OfficeIMO.Email;

internal static class MimeAddressParser {
    internal static IEnumerable<EmailAddress> ParseMany(string? input, bool allowSemicolonSeparator = false) {
        if (string.IsNullOrWhiteSpace(input)) yield break;
        foreach (string item in Split(input!, allowSemicolonSeparator)) {
            EmailAddress? address = ParseOne(item);
            if (address != null) yield return address;
        }
    }

    internal static EmailAddress? ParseOne(string? input) {
        if (string.IsNullOrWhiteSpace(input)) return null;
        string raw = input!.Trim();
        int less = raw.LastIndexOf('<');
        int greater = less >= 0 ? raw.IndexOf('>', less + 1) : -1;
        if (less >= 0 && greater > less) {
            string display = DecodeDisplayName(raw.Substring(0, less));
            string address = raw.Substring(less + 1, greater - less - 1).Trim();
            return new EmailAddress(address.Length == 0 ? null : address, display.Length == 0 ? null : display, raw);
        }
        return new EmailAddress(raw, null, raw);
    }

    internal static EmailAddress? ParseOne(string? input, IList<EmailDiagnostic> diagnostics, string location) {
        if (string.IsNullOrWhiteSpace(input)) return null;
        return ParseOne(MimeTextCodec.DecodeHeader(input!, diagnostics, location));
    }

    internal static IEnumerable<EmailAddress> ParseMany(string? input, IList<EmailDiagnostic> diagnostics,
        string location, bool allowSemicolonSeparator = false) {
        if (string.IsNullOrWhiteSpace(input)) yield break;
        foreach (string item in Split(input!, allowSemicolonSeparator)) {
            EmailAddress? address = ParseOne(item, diagnostics, location);
            if (address != null) yield return address;
        }
    }

    private static IEnumerable<string> Split(string input, bool allowSemicolonSeparator) {
        StringBuilder current = new StringBuilder();
        bool quoted = false;
        bool escaped = false;
        bool inGroup = false;
        int angleDepth = 0;
        int commentDepth = 0;
        for (int index = 0; index < input.Length; index++) {
            char character = input[index];
            if (escaped) {
                current.Append(character);
                escaped = false;
                continue;
            }
            if (character == '\\' && (quoted || commentDepth > 0)) {
                current.Append(character);
                escaped = true;
                continue;
            }
            if (!quoted && commentDepth == 0 && character == '=' && index + 1 < input.Length && input[index + 1] == '?') {
                int encodedWordEnd = input.IndexOf("?=", index + 2, StringComparison.Ordinal);
                if (encodedWordEnd >= 0) {
                    current.Append(input, index, encodedWordEnd + 2 - index);
                    index = encodedWordEnd + 1;
                    continue;
                }
            }
            if (character == '"' && commentDepth == 0) quoted = !quoted;
            if (!quoted) {
                if (character == '<') angleDepth++;
                if (character == '>' && angleDepth > 0) angleDepth--;
                if (character == '(') commentDepth++;
                if (character == ')' && commentDepth > 0) commentDepth--;
            }
            if (character == ':' && !quoted && angleDepth == 0 && commentDepth == 0 && !inGroup &&
                IsGroupLabel(current)) {
                current.Clear();
                inGroup = true;
                continue;
            }
            if (character == ';' && !quoted && angleDepth == 0 && commentDepth == 0 && inGroup) {
                if (current.Length > 0) yield return current.ToString();
                current.Clear();
                inGroup = false;
                continue;
            }
            if ((character == ',' || (allowSemicolonSeparator && character == ';')) &&
                !quoted && angleDepth == 0 && commentDepth == 0) {
                if (current.Length > 0) yield return current.ToString();
                current.Clear();
            } else {
                current.Append(character);
            }
        }
        if (current.Length > 0) yield return current.ToString();
    }

    private static bool IsGroupLabel(StringBuilder value) {
        string candidate = value.ToString().Trim();
        return candidate.Length > 0 && candidate.IndexOf('@') < 0;
    }

    private static string DecodeDisplayName(string value) {
        string display = value.Trim();
        if (display.Length < 2 || display[0] != '"' || display[display.Length - 1] != '"') {
            return display.Trim('"');
        }

        var result = new StringBuilder(display.Length - 2);
        for (int index = 1; index < display.Length - 1; index++) {
            char character = display[index];
            if (character == '\\' && index + 1 < display.Length - 1) {
                result.Append(display[++index]);
            } else {
                result.Append(character);
            }
        }
        return result.ToString();
    }
}
