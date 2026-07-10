namespace OfficeIMO.AsciiDoc;

internal static class AsciiDocLineReader {
    internal static IReadOnlyList<AsciiDocSourceLine> Read(string source) {
        var lines = new List<AsciiDocSourceLine>();
        int offset = 0;
        int lineNumber = 1;
        while (offset < source.Length) {
            int start = offset;
            while (offset < source.Length && source[offset] != '\r' && source[offset] != '\n') offset++;
            int contentEnd = offset;
            if (offset < source.Length && source[offset] == '\r') {
                offset++;
                if (offset < source.Length && source[offset] == '\n') offset++;
            } else if (offset < source.Length && source[offset] == '\n') {
                offset++;
            }

            lines.Add(new AsciiDocSourceLine(source, lineNumber++, start, contentEnd, offset));
        }

        return lines;
    }
}
