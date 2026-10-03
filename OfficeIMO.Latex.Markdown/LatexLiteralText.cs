namespace OfficeIMO.Latex.Markdown;

/// <summary>Decodes the literal escapes emitted by the bridge without executing commands.</summary>
internal static class LatexLiteralText {
    internal static string Decode(string source, System.Threading.CancellationToken cancellationToken = default, bool preserveTilde = true) {
        var output = new StringBuilder(source.Length);
        for (int index = 0; index < source.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            char current = source[index];
            if (current == '%') {
                while (index + 1 < source.Length && source[index + 1] != '\r' && source[index + 1] != '\n') {
                    if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                    index++;
                }
                if (index + 1 < source.Length && source[index + 1] == '\r') index++;
                if (index + 1 < source.Length && source[index + 1] == '\n') index++;
                continue;
            }
            if (current == '{' || current == '}') continue;
            if (current != '\\' || index + 1 >= source.Length) { output.Append(current == '~' && !preserveTilde ? ' ' : current); continue; }
            char next = source[index + 1];
            if ("{}%$&#_\\".IndexOf(next) >= 0) { output.Append(next); index++; continue; }
            int end = index + 1;
            while (end < source.Length && char.IsLetter(source[end])) {
                if ((end & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                end++;
            }
            string command = source.Substring(index + 1, end - index - 1);
            string? literal = command == "textbackslash" ? "\\" : command == "textasciitilde" ? "~" : command == "textasciicircum" ? "^" : null;
            if (literal == null) { output.Append(current); continue; }
            output.Append(literal);
            index = end - 1;
            if (end + 1 < source.Length && source[end] == '{' && source[end + 1] == '}') index += 2;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return output.ToString();
    }
}
