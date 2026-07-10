namespace OfficeIMO.Email;

internal static class MimeAddressParser {
    internal static IEnumerable<EmailAddress> ParseMany(string? input) {
        if (string.IsNullOrWhiteSpace(input)) yield break;
        foreach (string item in Split(input!)) {
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
            string display = raw.Substring(0, less).Trim().Trim('"');
            string address = raw.Substring(less + 1, greater - less - 1).Trim();
            return new EmailAddress(address.Length == 0 ? null : address, display.Length == 0 ? null : display, raw);
        }
        return new EmailAddress(raw, null, raw);
    }

    private static IEnumerable<string> Split(string input) {
        StringBuilder current = new StringBuilder();
        bool quoted = false;
        int angleDepth = 0;
        int commentDepth = 0;
        foreach (char character in input) {
            if (character == '"' && commentDepth == 0) quoted = !quoted;
            if (!quoted) {
                if (character == '<') angleDepth++;
                if (character == '>' && angleDepth > 0) angleDepth--;
                if (character == '(') commentDepth++;
                if (character == ')' && commentDepth > 0) commentDepth--;
            }
            if (character == ',' && !quoted && angleDepth == 0 && commentDepth == 0) {
                if (current.Length > 0) yield return current.ToString();
                current.Clear();
            } else {
                current.Append(character);
            }
        }
        if (current.Length > 0) yield return current.ToString();
    }
}
