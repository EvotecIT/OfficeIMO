namespace OfficeIMO.Pdf;

internal static class PdfInlineImageDataScanner {
    public static int FindLength(string content, int dataStart) {
        int index = dataStart;
        while (index + 2 < content.Length) {
            if (char.IsWhiteSpace(content[index]) &&
                content[index + 1] == 'E' &&
                content[index + 2] == 'I' &&
                (index + 3 >= content.Length || IsDelimiter(content[index + 3])) &&
                IsLikelyContentAfterInlineImage(content, index + 3)) {
                return index - dataStart;
            }

            index++;
        }

        return -1;
    }

    public static bool IsTerminatorAt(string content, int index) =>
        index + 1 < content.Length &&
        content[index] == 'E' &&
        content[index + 1] == 'I' &&
        (index + 2 >= content.Length || IsDelimiter(content[index + 2])) &&
        IsLikelyContentAfterInlineImage(content, index + 2);

    private static bool IsLikelyContentAfterInlineImage(string content, int index) {
        SkipWhitespaceAndComments(content, ref index);
        if (index >= content.Length) {
            return true;
        }

        string token = ReadToken(content, index, out int afterToken);
        if (IsKnownContentOperator(token)) {
            return true;
        }

        if (!IsLikelyOperandStart(content[index])) {
            return false;
        }

        int scan = afterToken;
        int limit = Math.Min(content.Length, index + 128);
        while (scan < limit) {
            SkipWhitespaceAndComments(content, ref scan);
            if (scan >= limit) {
                return false;
            }

            if (scan + 1 < content.Length &&
                content[scan] == 'E' &&
                content[scan + 1] == 'I' &&
                (scan + 2 >= content.Length || IsDelimiter(content[scan + 2]))) {
                return false;
            }

            string next = ReadToken(content, scan, out int nextIndex);
            if (next.Length == 0) {
                scan++;
                continue;
            }

            if (IsKnownContentOperator(next)) {
                return true;
            }

            scan = nextIndex;
        }

        return false;
    }

    private static bool IsLikelyOperandStart(char ch) =>
        ch == '/' ||
        ch == '[' ||
        ch == '<' ||
        ch == '(' ||
        ch == '+' ||
        ch == '-' ||
        ch == '.' ||
        char.IsDigit(ch);

    private static string ReadToken(string content, int index, out int nextIndex) {
        nextIndex = index;
        if (index >= content.Length) {
            return string.Empty;
        }

        char ch = content[index];
        if (ch == '\'' || ch == '"') {
            nextIndex = index + 1;
            return ch.ToString();
        }

        int start = index;
        while (nextIndex < content.Length && !IsDelimiter(content[nextIndex])) {
            nextIndex++;
        }

        return content.Substring(start, nextIndex - start);
    }

    private static void SkipWhitespaceAndComments(string content, ref int index) {
        while (index < content.Length) {
            while (index < content.Length && char.IsWhiteSpace(content[index])) {
                index++;
            }

            if (index >= content.Length || content[index] != '%') {
                return;
            }

            while (index < content.Length && content[index] != '\r' && content[index] != '\n') {
                index++;
            }
        }
    }

    private static bool IsKnownContentOperator(string token) {
        switch (token) {
            case "q":
            case "Q":
            case "cm":
            case "w":
            case "J":
            case "j":
            case "M":
            case "d":
            case "ri":
            case "i":
            case "gs":
            case "m":
            case "l":
            case "c":
            case "v":
            case "y":
            case "h":
            case "re":
            case "S":
            case "s":
            case "f":
            case "F":
            case "f*":
            case "B":
            case "B*":
            case "b":
            case "b*":
            case "n":
            case "W":
            case "W*":
            case "BT":
            case "BI":
            case "ET":
            case "Tc":
            case "Tw":
            case "Tz":
            case "TL":
            case "Tf":
            case "Tr":
            case "Ts":
            case "Td":
            case "TD":
            case "Tm":
            case "T*":
            case "Tj":
            case "TJ":
            case "'":
            case "\"":
            case "d0":
            case "d1":
            case "CS":
            case "cs":
            case "SC":
            case "SCN":
            case "sc":
            case "scn":
            case "G":
            case "g":
            case "RG":
            case "rg":
            case "K":
            case "k":
            case "sh":
            case "Do":
            case "MP":
            case "DP":
            case "BMC":
            case "BDC":
            case "EMC":
            case "BX":
            case "EX":
                return true;
            default:
                return false;
        }
    }

    private static bool IsDelimiter(char ch) =>
        char.IsWhiteSpace(ch) ||
        ch == '/' ||
        ch == '[' ||
        ch == ']' ||
        ch == '<' ||
        ch == '>' ||
        ch == '(' ||
        ch == ')' ||
        ch == '%';
}
