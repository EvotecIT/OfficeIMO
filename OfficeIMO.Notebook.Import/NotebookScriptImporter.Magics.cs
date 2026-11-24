namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookScriptImporter {
    private static string UnescapePercentCode(string line, string language, bool commentMagics, bool explicitCell,
        ref bool magicContinues) {
        ReadOnlySpan<char> text = line.AsSpan().TrimStart();
        int indent = line.Length - text.Length;
        ReadOnlySpan<char> candidate = RemoveComment(text);
        ReadOnlySpan<char> marker = candidate;
        while (!marker.IsEmpty && marker[0] == '#') marker = RemoveComment(marker);

        bool force = marker.EndsWith("# escape".AsSpan(), StringComparison.Ordinal);
        bool keep = marker.EndsWith("# noescape".AsSpan(), StringComparison.Ordinal);
        bool isMagic = magicContinues || ((!keep && commentMagics) || force) &&
            (IsPercentMagic(marker) || language == "python" && IsPythonMagic(marker, explicitCell));
        magicContinues = isMagic && language == "python" && marker.TrimEnd().EndsWith("\\".AsSpan(), StringComparison.Ordinal);
        return isMagic && candidate.Length != text.Length ? line.Substring(0, indent) + candidate.ToString() : line;
    }

    private static bool IsPercentMagic(ReadOnlySpan<char> text) {
        int prefix = 0;
        while (prefix < text.Length && prefix < 3 && text[prefix] == '%') prefix++;
        return prefix > 0 && prefix < text.Length && IsAsciiLetter(text[prefix]);
    }

    private static bool IsPythonMagic(ReadOnlySpan<char> text, bool explicitCell) {
        if (text.IsEmpty) return false;
        if (text[0] is '!' or '?') {
            ReadOnlySpan<char> command = text.Slice(1).TrimStart();
            return !command.IsEmpty && (IsAsciiLetter(command[0]) || command[0] is '.' or '~' or '$' or '\\' or '/' or '{' or '}');
        }

        int wordEnd = 0;
        while (wordEnd < text.Length && !char.IsWhiteSpace(text[wordEnd]) && text[wordEnd] != '=') wordEnd++;
        ReadOnlySpan<char> tail = text.Slice(wordEnd).TrimStart();
        if (tail.Length > 0 && tail[0] == '=' && IsPythonName(text.Slice(0, wordEnd))) {
            ReadOnlySpan<char> assignment = tail.Slice(1).TrimStart();
            return IsPercentMagic(assignment) || assignment.Length > 1 && assignment[0] == '!' && IsAsciiLetter(assignment[1]);
        }

        ReadOnlySpan<char> token = text.Slice(0, wordEnd);
        if (explicitCell && tail.IsEmpty && token.Length > 0 && token[token.Length - 1] == '?') return true;
        if (!tail.IsEmpty && tail[0] is '=' or ',') return false;

        foreach (string command in PythonMagicCommands) {
            if (token.SequenceEqual(command.AsSpan())) return true;
        }
        return false;
    }

    private static readonly string[] PythonMagicCommands = [
        "cat", "cd", "cp", "mv", "rm", "rmdir", "mkdir", "copy", "ddir", "echo", "ls", "ldir", "ren"
    ];

    private static bool IsPythonName(ReadOnlySpan<char> word) {
        if (word.IsEmpty || !IsAsciiLetter(word[0]) && word[0] != '_') return false;
        foreach (char c in word) {
            if (!IsAsciiLetter(c) && c is not (>= '0' and <= '9' or '_' or '$')) return false;
        }
        return true;
    }

    private static bool IsAsciiLetter(char c) => c is >= 'A' and <= 'Z' or >= 'a' and <= 'z';
}
