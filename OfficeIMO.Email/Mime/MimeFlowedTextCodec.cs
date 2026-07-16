namespace OfficeIMO.Email;

/// <summary>
/// Decodes RFC 3676 format=flowed plain text into the message model's logical text representation.
/// </summary>
internal static class MimeFlowedTextCodec {
    internal static string Decode(string value, bool deleteSpace) {
        if (string.IsNullOrEmpty(value)) return value;

        string newline = DetectNewline(value);
        string normalized = value.Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        StringBuilder output = new StringBuilder(value.Length);

        for (int index = 0; index < lines.Length; index++) {
            FlowedLine current = Parse(lines[index]);
            output.Append(current.QuotePrefix);

            while (IsFlowed(current) && index + 1 < lines.Length) {
                FlowedLine next = Parse(lines[index + 1]);
                if (next.QuoteDepth != current.QuoteDepth) break;

                output.Append(deleteSpace
                    ? current.Content.Substring(0, current.Content.Length - 1)
                    : current.Content);
                current = next;
                index++;
            }

            if (deleteSpace && IsFlowed(current) && index + 1 == lines.Length) {
                output.Append(current.Content.Substring(0, current.Content.Length - 1));
            } else {
                output.Append(current.Content);
            }
            if (index + 1 < lines.Length) output.Append(newline);
        }

        return output.ToString();
    }

    private static FlowedLine Parse(string physicalLine) {
        string line = physicalLine.Length > 0 && physicalLine[0] == ' '
            ? physicalLine.Substring(1)
            : physicalLine;
        int quoteDepth = 0;
        while (quoteDepth < line.Length && line[quoteDepth] == '>') quoteDepth++;
        return new FlowedLine(line.Substring(0, quoteDepth), line.Substring(quoteDepth), quoteDepth);
    }

    private static bool IsFlowed(FlowedLine line) => line.Content.EndsWith(" ", StringComparison.Ordinal) &&
        !string.Equals(line.Content, "-- ", StringComparison.Ordinal);

    private static string DetectNewline(string value) {
        if (value.IndexOf("\r\n", StringComparison.Ordinal) >= 0) return "\r\n";
        if (value.IndexOf('\r') >= 0) return "\r";
        return "\n";
    }

    private readonly struct FlowedLine {
        internal FlowedLine(string quotePrefix, string content, int quoteDepth) {
            QuotePrefix = quotePrefix;
            Content = content;
            QuoteDepth = quoteDepth;
        }

        internal string QuotePrefix { get; }
        internal string Content { get; }
        internal int QuoteDepth { get; }
    }
}
