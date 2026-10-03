namespace OfficeIMO.Bibliography;

/// <summary>Scans BibTeX name words without splitting brace-protected components.</summary>
internal static class BibNameSyntax {
    internal static string[] Words(string value, CancellationToken cancellationToken) {
        var words = new List<string>();
        int start = -1;
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char character = value[index];
            if (depth == 0 && char.IsWhiteSpace(character)) {
                if (start >= 0) { words.Add(value.Substring(start, index - start)); start = -1; }
                continue;
            }
            if (start < 0) start = index;
            if (character == '\\' && index + 1 < value.Length) { index++; continue; }
            if (character == '{') depth++;
            else if (character == '}' && depth > 0) depth--;
        }
        if (start >= 0) words.Add(value.Substring(start));
        cancellationToken.ThrowIfCancellationRequested();
        return words.ToArray();
    }

    internal static bool IsOuterGroup(string value) {
        if (value.Length < 2 || value[0] != '{' || value[value.Length - 1] != '}') return false;
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\\' && index + 1 < value.Length) { index++; continue; }
            if (value[index] == '{') depth++;
            else if (value[index] == '}' && --depth == 0) return index == value.Length - 1;
        }
        return false;
    }

    internal static string Unprotect(string value) {
        var result = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character == '\\' && index + 1 < value.Length) { result.Append(character).Append(value[++index]); continue; }
            if (character != '{' && character != '}') result.Append(character);
        }
        return result.ToString();
    }

    internal static bool StartsWithLowercaseLetter(string value) {
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character == '{') { depth++; continue; }
            if (character == '}') { depth--; continue; }
            if (depth == 0 && char.IsLetter(character)) return char.IsLower(character);
        }
        return false;
    }
}
