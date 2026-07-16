namespace OfficeIMO.Email.AddressBook;

internal static class OabSearchText {
    internal static string Normalize(string? value, int maximumCharacters) {
        if (string.IsNullOrEmpty(value) || maximumCharacters <= 0) return string.Empty;
        var builder = new StringBuilder(Math.Min(value!.Length, maximumCharacters));
        bool pendingSpace = false;
        for (int index = 0; index < value.Length && builder.Length < maximumCharacters; index++) {
            char character = value[index];
            if (char.IsWhiteSpace(character) || char.IsControl(character)) {
                pendingSpace = builder.Length > 0;
                continue;
            }
            if (pendingSpace && builder.Length < maximumCharacters) builder.Append(' ');
            pendingSpace = false;
            if (builder.Length < maximumCharacters) builder.Append(character);
        }
        return builder.ToString();
    }

    internal static string CreateSnippet(string text, IReadOnlyList<string> terms, int maximumCharacters) {
        int matchIndex = int.MaxValue;
        int matchLength = 0;
        foreach (string term in terms) {
            int found = text.IndexOf(term, StringComparison.OrdinalIgnoreCase);
            if (found >= 0 && found < matchIndex) {
                matchIndex = found;
                matchLength = term.Length;
            }
        }
        if (matchIndex == int.MaxValue) return text.Length <= maximumCharacters
            ? text
            : text.Substring(0, maximumCharacters) + "…";
        int context = Math.Max(0, (maximumCharacters - matchLength) / 2);
        int start = Math.Max(0, matchIndex - context);
        int length = Math.Min(maximumCharacters, text.Length - start);
        string snippet = text.Substring(start, length);
        if (start > 0) snippet = "…" + snippet;
        if (start + length < text.Length) snippet += "…";
        return snippet;
    }
}
