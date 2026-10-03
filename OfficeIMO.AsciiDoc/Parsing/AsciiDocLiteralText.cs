namespace OfficeIMO.AsciiDoc;

/// <summary>Protects literal paragraph lines from the native block classifier.</summary>
internal static class AsciiDocLiteralText {
    internal static string EscapeBlockStarts(string value) => Transform(value, true, true);
    internal static string UnescapeBlockStarts(string value, bool startsAtLineBeginning) => Transform(value, false, startsAtLineBeginning);

    private static string Transform(string value, bool escape, bool atLineBeginning) {
        var output = new StringBuilder(value.Length);
        int start = 0;
        while (start < value.Length) {
            int end = start;
            while (end < value.Length && value[end] != '\r' && value[end] != '\n') end++;
            string line = value.Substring(start, end - start);
            if (atLineBeginning && line.Length > 0) {
                if (escape && AsciiDocLineClassifier.TryParseDescriptionListItem(line, out var description))
                    line = line.Insert(description.MarkerStart, "\\");
                else if (escape && !AsciiDocLineClassifier.IsBlank(line) && AsciiDocLineClassifier.IsBlockStart(line)) output.Append('\\');
                else if (!escape && line[0] == '\\' && line.Length > 1 && AsciiDocLineClassifier.IsBlockStart(line.Substring(1))) line = line.Substring(1);
            }
            output.Append(line);
            if (end < value.Length) {
                output.Append(value[end++]);
                if (end < value.Length && value[end - 1] == '\r' && value[end] == '\n') output.Append(value[end++]);
                atLineBeginning = true;
            }
            start = end;
        }
        return output.ToString();
    }
}
