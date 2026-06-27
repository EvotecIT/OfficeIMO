using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static Dictionary<string, object?> ParseFrontMatter(string[] lines, int start, int end) {
        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in ParseFrontMatterEntries(lines, start, end, null)) {
            dict[entry.Key] = entry.Value;
        }

        return dict;
    }

    private static FrontMatterBlock ParseFrontMatterBlock(string[] lines, int start, int end, MarkdownReaderState state) =>
        FrontMatterBlock.FromEntries(ParseFrontMatterEntries(lines, start, end, state));

    private static IReadOnlyList<FrontMatterBlock.Entry> ParseFrontMatterEntries(string[] lines, int start, int end, MarkdownReaderState? state) {
        var entries = new Dictionary<string, FrontMatterBlock.Entry>(StringComparer.OrdinalIgnoreCase);
        for (int i = start; i <= end; i++) {
            var line = lines[i]; if (string.IsNullOrWhiteSpace(line)) continue;
            int idx = line.IndexOf(':'); if (idx <= 0) continue;
            string key = line.Substring(0, idx).Trim(); string val = line.Substring(idx + 1).TrimStart();
            if (key.Length == 0) continue;

            var keySpan = CreateFrontMatterKeySpan(lines, i, idx, state);
            var valueSpan = CreateFrontMatterValueSpan(line, i, idx, val, state);
            if (val == "|") {
                var sb = new StringBuilder(); int j = i + 1;
                while (j <= end) { var raw = lines[j]; if (raw.StartsWith("  ")) { sb.AppendLine(raw.Substring(2)); j++; } else break; }
                valueSpan = CreateFrontMatterLiteralValueSpan(lines, i + 1, j - 1, state);
                i = j - 1; entries[key] = new FrontMatterBlock.Entry(key, sb.ToString().TrimEnd(), keySpan, valueSpan); continue;
            }
            if (val.StartsWith("[") && val.EndsWith("]")) {
                var inner = val.Substring(1, val.Length - 2).Trim(); var items = new List<string>(); var token = new StringBuilder(); bool inQuotes = false;
                for (int k = 0; k < inner.Length; k++) { char ch = inner[k]; if (ch == '\"') { inQuotes = !inQuotes; continue; } if (ch == ',' && !inQuotes) { items.Add(token.ToString().Trim()); token.Clear(); continue; } token.Append(ch); }
                if (token.Length > 0) items.Add(token.ToString().Trim());
                entries[key] = new FrontMatterBlock.Entry(key, items, keySpan, valueSpan);
            } else if (string.Equals(val, "true", StringComparison.OrdinalIgnoreCase)) { entries[key] = new FrontMatterBlock.Entry(key, true, keySpan, valueSpan); } else if (string.Equals(val, "false", StringComparison.OrdinalIgnoreCase)) { entries[key] = new FrontMatterBlock.Entry(key, false, keySpan, valueSpan); } else if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out var num)) { entries[key] = new FrontMatterBlock.Entry(key, num, keySpan, valueSpan); } else if (val.StartsWith("\"") && val.EndsWith("\"")) { entries[key] = new FrontMatterBlock.Entry(key, val.Length >= 2 ? val.Substring(1, val.Length - 2) : string.Empty, keySpan, valueSpan); } else { entries[key] = new FrontMatterBlock.Entry(key, val, keySpan, valueSpan); }
        }

        return entries.Values.ToArray();
    }

    private static MarkdownSourceSpan? CreateFrontMatterKeySpan(string[] lines, int lineIndex, int colonIndex, MarkdownReaderState? state) {
        var line = lines[lineIndex];
        int keyStart = 0;
        while (keyStart < colonIndex && char.IsWhiteSpace(line[keyStart])) keyStart++;
        int keyEnd = colonIndex - 1;
        while (keyEnd >= keyStart && char.IsWhiteSpace(line[keyEnd])) keyEnd--;
        if (keyEnd < keyStart) {
            return null;
        }

        int absoluteLine = (state?.SourceLineOffset ?? 0) + lineIndex + 1;
        return CreateSpan(state, absoluteLine, keyStart + 1, absoluteLine, keyEnd + 1);
    }

    private static MarkdownSourceSpan? CreateFrontMatterValueSpan(string line, int lineIndex, int colonIndex, string value, MarkdownReaderState? state) {
        if (string.IsNullOrEmpty(value)) {
            return null;
        }

        int valueStart = colonIndex + 1;
        while (valueStart < line.Length && char.IsWhiteSpace(line[valueStart])) valueStart++;
        if (valueStart >= line.Length) {
            return null;
        }

        int absoluteLine = (state?.SourceLineOffset ?? 0) + lineIndex + 1;
        return CreateSpan(state, absoluteLine, valueStart + 1, absoluteLine, valueStart + value.Length);
    }

    private static MarkdownSourceSpan? CreateFrontMatterLiteralValueSpan(string[] lines, int startLineIndex, int endLineIndex, MarkdownReaderState? state) {
        if (endLineIndex < startLineIndex) {
            return null;
        }

        int absoluteStartLine = (state?.SourceLineOffset ?? 0) + startLineIndex + 1;
        int absoluteEndLine = (state?.SourceLineOffset ?? 0) + endLineIndex + 1;
        int endColumn = Math.Max(3, lines[endLineIndex].Length);
        return CreateSpan(state, absoluteStartLine, 3, absoluteEndLine, endColumn);
    }
}
