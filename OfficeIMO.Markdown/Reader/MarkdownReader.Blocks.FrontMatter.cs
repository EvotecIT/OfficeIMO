using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static Dictionary<string, object?> ParseFrontMatter(string[] lines, int start, int end) {
        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        for (int i = start; i <= end; i++) {
            var line = lines[i]; if (string.IsNullOrWhiteSpace(line)) continue;
            int idx = line.IndexOf(':'); if (idx <= 0) continue;
            string key = line.Substring(0, idx).Trim(); string val = line.Substring(idx + 1).TrimStart();
            if (val == "|") {
                var sb = new StringBuilder(); int j = i + 1;
                while (j <= end) { var raw = lines[j]; if (raw.StartsWith("  ")) { sb.AppendLine(raw.Substring(2)); j++; } else break; }
                i = j - 1; dict[key] = sb.ToString().TrimEnd(); continue;
            }
            if (val.StartsWith("[") && val.EndsWith("]")) {
                var inner = val.Substring(1, val.Length - 2).Trim(); var items = new List<string>(); var token = new StringBuilder(); bool inQuotes = false;
                for (int k = 0; k < inner.Length; k++) { char ch = inner[k]; if (ch == '\"') { inQuotes = !inQuotes; continue; } if (ch == ',' && !inQuotes) { items.Add(token.ToString().Trim()); token.Clear(); continue; } token.Append(ch); }
                if (token.Length > 0) items.Add(token.ToString().Trim());
                dict[key] = items;
            } else if (string.Equals(val, "true", StringComparison.OrdinalIgnoreCase)) { dict[key] = true; } else if (string.Equals(val, "false", StringComparison.OrdinalIgnoreCase)) { dict[key] = false; } else if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out var num)) { dict[key] = num; } else if (val.StartsWith("\"") && val.EndsWith("\"")) { dict[key] = val.Length >= 2 ? val.Substring(1, val.Length - 2) : string.Empty; } else { dict[key] = val; }
        }
        return dict;
    }
}
