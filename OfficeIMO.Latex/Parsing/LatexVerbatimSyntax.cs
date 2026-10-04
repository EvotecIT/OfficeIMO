namespace OfficeIMO.Latex;

/// <summary>Parses the bounded delimiter syntax used by opaque LaTeX environments.</summary>
internal static class LatexVerbatimSyntax {
    internal static bool TryReadEnvironmentOpening(
        string source,
        int start,
        out string environmentName,
        out int contentStart, System.Threading.CancellationToken cancellationToken = default) {
        environmentName = string.Empty;
        contentStart = start;
        if (!StartsWithControlWord(source, start, "begin")) return false;

        int cursor = start + 6;
        SkipArgumentTrivia(source, ref cursor, cancellationToken);
        if (cursor >= source.Length || source[cursor] != '{') return false;
        int nameStart = ++cursor;
        int nameEnd = Find(source, "}", nameStart, cancellationToken);
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
        out int closingEnd, System.Threading.CancellationToken cancellationToken = default) {
        closingStart = -1;
        closingEnd = -1;
        int candidate = searchStart;
        while (candidate < source.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            candidate = Find(source, "\\end", candidate, cancellationToken);
            if (candidate < 0) return false;
            if (TryReadEnvironmentName(source, candidate, "end", out string name, out int end, cancellationToken)
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
        out int end, System.Threading.CancellationToken cancellationToken) {
        environmentName = string.Empty;
        end = start;
        if (!StartsWithControlWord(source, start, controlWord)) return false;
        int cursor = start + controlWord.Length + 1;
        SkipArgumentTrivia(source, ref cursor, cancellationToken);
        if (cursor >= source.Length || source[cursor] != '{') return false;
        int nameStart = ++cursor;
        int nameEnd = Find(source, "}", nameStart, cancellationToken);
        if (nameEnd < 0) return false;
        environmentName = source.Substring(nameStart, nameEnd - nameStart).Trim();
        end = nameEnd + 1;
        return environmentName.Length > 0;
    }

    private static void SkipArgumentTrivia(string source, ref int cursor, System.Threading.CancellationToken cancellationToken) {
        while (cursor < source.Length) {
            if ((cursor & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (char.IsWhiteSpace(source[cursor])) {
                cursor++;
                continue;
            }
            if (source[cursor] != '%') return;
            cursor++;
            while (cursor < source.Length && source[cursor] != '\r' && source[cursor] != '\n') {
                if ((cursor & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                cursor++;
            }
        }
    }

    // Bounded search chunks let cancellable projection stop within a large opaque body.
    private static int Find(string source, string value, int start, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled) return source.IndexOf(value, start, StringComparison.Ordinal);
        while (start < source.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = Math.Min(4096, source.Length - start);
            int found = source.IndexOf(value, start, length, StringComparison.Ordinal);
            if (found >= 0) return found;
            if (length < 4096) return -1;
            start += length - value.Length + 1;
        }
        return -1;
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
