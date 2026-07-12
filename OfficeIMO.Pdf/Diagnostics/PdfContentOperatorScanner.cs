namespace OfficeIMO.Pdf;

internal static class PdfContentOperatorScanner {
    private static readonly HashSet<string> Operators = new(StringComparer.Ordinal) {
        "b", "B", "b*", "B*", "BDC", "BI", "BMC", "BT", "BX",
        "c", "cm", "CS", "cs", "d", "d0", "d1", "Do", "DP",
        "EI", "EMC", "ET", "EX", "f", "F", "f*", "G", "g", "gs",
        "h", "i", "ID", "j", "J", "K", "k", "l", "m", "M", "MP",
        "n", "q", "Q", "re", "RG", "rg", "ri", "s", "S", "SC", "sc",
        "SCN", "scn", "sh", "T*", "Tc", "Td", "TD", "Tf", "Tj", "TJ",
        "TL", "Tm", "Tr", "Ts", "Tw", "Tz", "v", "w", "W", "W*", "y", "'", "\""
    };

    internal static void AppendOperators(string content, List<string> destination, int maximum, ref bool truncated) {
        int index = 0;
        while (index < content.Length) {
            SkipWhitespaceAndComments(content, ref index);
            if (index >= content.Length) break;
            char current = content[index];
            if (current == '(') {
                SkipLiteralString(content, ref index);
                continue;
            }

            if (current == '<' && index + 1 < content.Length && content[index + 1] != '<') {
                index = Math.Min(content.Length, content.IndexOf('>', index + 1) is int end && end >= 0 ? end + 1 : content.Length);
                continue;
            }

            string token = ReadToken(content, ref index);
            if (!Operators.Contains(token)) continue;
            if (destination.Count >= maximum) {
                truncated = true;
                return;
            }

            destination.Add(token);
            if (token == "ID") {
                int dataStart = index;
                if (dataStart < content.Length && char.IsWhiteSpace(content[dataStart])) dataStart++;
                int length = PdfInlineImageDataScanner.FindLength(content, dataStart);
                if (length >= 0) index = dataStart + length;
            }
        }
    }

    private static void SkipWhitespaceAndComments(string content, ref int index) {
        while (index < content.Length) {
            if (char.IsWhiteSpace(content[index])) {
                index++;
            } else if (content[index] == '%') {
                while (index < content.Length && content[index] != '\r' && content[index] != '\n') index++;
            } else {
                break;
            }
        }
    }

    private static void SkipLiteralString(string content, ref int index) {
        int depth = 0;
        bool escaped = false;
        while (index < content.Length) {
            char value = content[index++];
            if (escaped) {
                escaped = false;
            } else if (value == '\\') {
                escaped = true;
            } else if (value == '(') {
                depth++;
            } else if (value == ')' && --depth == 0) {
                return;
            }
        }
    }

    private static string ReadToken(string content, ref int index) {
        char first = content[index];
        if (first == '\'' || first == '"') {
            index++;
            return first.ToString();
        }

        int start = index;
        if (IsDelimiter(first)) {
            index++;
            if (index < content.Length && content[index] == first && (first == '<' || first == '>')) index++;
        } else {
            while (index < content.Length && !char.IsWhiteSpace(content[index]) && !IsDelimiter(content[index])) index++;
        }

        return content.Substring(start, index - start);
    }

    private static bool IsDelimiter(char value) =>
        value == '(' || value == ')' || value == '<' || value == '>' || value == '[' || value == ']' ||
        value == '{' || value == '}' || value == '/' || value == '%';
}
