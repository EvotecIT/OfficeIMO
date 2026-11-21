namespace OfficeIMO.Yaml;

/// <summary>Formats the source line around a position with a caret under the column, for error messages and diagnostics.</summary>
internal static class YamlExcerpt {
    /// <summary>The widest excerpt shown; longer lines are cut around the column.</summary>
    private const int Width = 100;

    /// <summary>Returns a line break, the line holding <paramref name="offset"/> and a caret line, or an empty string.</summary>
    public static string Format(string source, int offset) {
        if (offset < 0 || offset > source.Length) return string.Empty;
        int start = offset;
        while (start > 0 && source[start - 1] != '\n' && source[start - 1] != '\r') start--;
        int end = offset;
        while (end < source.Length && source[end] != '\n' && source[end] != '\r') end++;
        int from = start, to = end;
        if (to - from > Width) {
            from = Math.Max(start, offset - Width / 2);
            to = Math.Min(end, from + Width);
        }
        var text = new StringBuilder(Environment.NewLine.Length * 2 + (to - from) * 2 + 8);
        text.Append(Environment.NewLine).Append("  ");
        for (int i = from; i < to; i++) text.Append(source[i] == '\t' ? ' ' : source[i]);
        text.Append(Environment.NewLine).Append("  ").Append(' ', offset - from).Append('^');
        return text.ToString();
    }
}
