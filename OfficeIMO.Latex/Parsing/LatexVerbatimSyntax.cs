namespace OfficeIMO.Latex;

/// <summary>Parses the bounded delimiter syntax used by opaque LaTeX environments.</summary>
internal static class LatexVerbatimSyntax {
    internal static bool TryReadEnvironmentOpening(
        string source,
        int start,
        out string environmentName,
        out int contentStart) {
        environmentName = string.Empty;
        contentStart = start;
        if (!StartsWithControlWord(source, start, "begin")) return false;

        int cursor = start + 6;
        SkipArgumentTrivia(source, ref cursor);
        if (cursor >= source.Length || source[cursor] != '{') return false;
        int nameStart = ++cursor;
        int nameEnd = source.IndexOf('}', nameStart);
        if (nameEnd < 0) return false;
        environmentName = source.Substring(nameStart, nameEnd - nameStart).Trim();
        if (environmentName.Length == 0) return false;
        contentStart = nameEnd + 1;
        return true;
    }

    internal static bool TryFindEnvironmentClosing(
        string source,
        int searchStart,
        string environmentName,
        out int closingStart,
        out int closingEnd) {
        closingStart = -1;
        closingEnd = -1;
        int candidate = searchStart;
        while (candidate < source.Length) {
            candidate = source.IndexOf("\\end", candidate, StringComparison.Ordinal);
            if (candidate < 0) return false;
            if (TryReadEnvironmentName(source, candidate, "end", out string name, out int end)
                && string.Equals(name, environmentName, StringComparison.Ordinal)) {
                closingStart = candidate;
                closingEnd = end;
                return true;
            }
            candidate++;
        }
        return false;
    }

    private static bool TryReadEnvironmentName(
        string source,
        int start,
        string controlWord,
        out string environmentName,
        out int end) {
        environmentName = string.Empty;
        end = start;
        if (!StartsWithControlWord(source, start, controlWord)) return false;
        int cursor = start + controlWord.Length + 1;
        SkipArgumentTrivia(source, ref cursor);
        if (cursor >= source.Length || source[cursor] != '{') return false;
        int nameStart = ++cursor;
        int nameEnd = source.IndexOf('}', nameStart);
        if (nameEnd < 0) return false;
        environmentName = source.Substring(nameStart, nameEnd - nameStart).Trim();
        end = nameEnd + 1;
        return environmentName.Length > 0;
    }

    private static void SkipArgumentTrivia(string source, ref int cursor) {
        while (cursor < source.Length) {
            if (char.IsWhiteSpace(source[cursor])) {
                cursor++;
                continue;
            }
            if (source[cursor] != '%') return;
            cursor++;
            while (cursor < source.Length && source[cursor] != '\r' && source[cursor] != '\n') cursor++;
        }
    }

    private static bool StartsWithControlWord(string source, int start, string name) {
        if (start + name.Length + 1 > source.Length || source[start] != '\\'
            || string.Compare(source, start + 1, name, 0, name.Length, StringComparison.Ordinal) != 0) {
            return false;
        }
        int end = start + name.Length + 1;
        return end >= source.Length || !IsControlWordCharacter(source[end]);
    }

    private static bool IsControlWordCharacter(char value) =>
        (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z') || value == '@';
}
