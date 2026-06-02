using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static int CountLeadingSpaces(string line) {
        if (string.IsNullOrEmpty(line)) return 0;
        int i = 0;
        while (i < line.Length && line[i] == ' ') i++;
        return i;
    }

    private static int CountLeadingIndentColumns(string line) {
        if (string.IsNullOrEmpty(line)) return 0;

        int columns = 0;
        for (int i = 0; i < line.Length; i++) {
            char ch = line[i];
            if (ch == ' ') {
                columns++;
                continue;
            }

            if (ch == '\t') {
                columns += 4 - (columns % 4);
                continue;
            }

            break;
        }

        return columns;
    }

    private static string StripLeadingIndentColumns(string line, int requiredColumns) {
        if (string.IsNullOrEmpty(line) || requiredColumns <= 0) return line ?? string.Empty;

        int columns = 0;
        int index = 0;
        while (index < line.Length && columns < requiredColumns) {
            char ch = line[index];
            if (ch == ' ') {
                columns++;
                index++;
                continue;
            }

            if (ch == '\t') {
                columns += 4 - (columns % 4);
                index++;
                continue;
            }

            break;
        }

        return line.Substring(index);
    }

}
