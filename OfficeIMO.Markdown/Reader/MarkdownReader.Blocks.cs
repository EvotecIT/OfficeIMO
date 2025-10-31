using System.Globalization;

namespace OfficeIMO.Markdown;

/// <summary>
/// Block parsing helpers for <see cref="MarkdownReader"/>.
/// </summary>
public static partial class MarkdownReader {
    private static bool IsAtxHeading(string line, out int level, out string text) {
        level = 0; text = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        int i = 0; while (i < line.Length && line[i] == '#') i++;
        if (i < 1 || i > 6) return false;
        if (i < line.Length && line[i] == ' ') { level = i; text = line.Substring(i + 1); return true; }
        return false;
    }

    private static bool IsCodeFenceOpen(string line, out string language) {
        language = string.Empty;
        if (line is null) return false;
        line = line.Trim();
        if (line.StartsWith("```")) { language = line.Length > 3 ? line.Substring(3).Trim() : string.Empty; return true; }
        return false;
    }
    private static bool IsCodeFenceClose(string line) => line.Trim() == "```";

    private static bool TryParseCaption(string line, out string caption) {
        caption = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.Trim();
        if (t.Length >= 3 && t[0] == '_' && t[t.Length - 1] == '_' && t.IndexOf('_', 1) == t.Length - 1) { caption = t.Substring(1, t.Length - 2); return true; }
        return false;
    }

    private static bool IsImageLine(string line) => TryParseImage(line, out _);
    private static bool TryParseImage(string line, out ImageBlock image) {
        image = null!;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.Trim();
        if (!t.StartsWith("![")) return false;
        int altEnd = t.IndexOf(']');
        if (altEnd < 2) return false;
        if (altEnd + 1 >= t.Length || t[altEnd + 1] != '(') return false;
        int parenClose = t.IndexOf(')', altEnd + 2);
        if (parenClose <= altEnd + 2) return false;
        string alt = t.Substring(2, altEnd - 2);
        string inner = t.Substring(altEnd + 2, parenClose - (altEnd + 2));
        string src; string? title = null;
        int spaceIdx = inner.IndexOf(' ');
        if (spaceIdx < 0) { src = inner.Trim(); } else { src = inner.Substring(0, spaceIdx).Trim(); string rest = inner.Substring(spaceIdx).Trim(); if (rest.Length >= 2 && rest[0] == '"' && rest[rest.Length - 1] == '"') title = rest.Substring(1, rest.Length - 2); }
        image = new ImageBlock(src, alt, title);
        // Optional attribute list: {width=.. height=..}
        if (parenClose + 1 < t.Length) {
            var rest = t.Substring(parenClose + 1).Trim();
            if (rest.StartsWith("{")) {
                int close = rest.IndexOf('}');
                if (close > 0) {
                    var attrs = rest.Substring(1, close - 1).Trim();
                    foreach (var part in attrs.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                        int eq = part.IndexOf('=');
                        if (eq > 0) {
                            var key = part.Substring(0, eq).Trim();
                            var val = part.Substring(eq + 1).Trim();
                            if (double.TryParse(val, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out var num)) {
                                if (string.Equals(key, "width", StringComparison.OrdinalIgnoreCase)) image.Width = num;
                                else if (string.Equals(key, "height", StringComparison.OrdinalIgnoreCase)) image.Height = num;
                            }
                        }
                    }
                }
            }
        }
        return true;
    }

    private static bool LooksLikeTableRow(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.Trim();
        if (trimmed.Length < 3 || !trimmed.Contains('|')) return false;

        var cells = SplitTableRow(trimmed);
        if (cells.Count == 0) return false;

        bool hasOuterPipes = trimmed[0] == '|' || trimmed[trimmed.Length - 1] == '|';
        if (!hasOuterPipes && cells.Count < 2) return false;

        return true;
    }

    private static TableBlock ParseTable(string[] lines, int start, int end) {
        var cells0 = SplitTableRow(lines[start]);
        var table = new TableBlock();
        if (start + 1 <= end && IsAlignmentRow(lines[start + 1])) {
            table.Headers.AddRange(cells0);
            var aligns = SplitTableRow(lines[start + 1]);
            for (int i = 0; i < aligns.Count; i++) table.Alignments.Add(ParseAlignmentCell(aligns[i]));
            for (int i = start + 2; i <= end; i++) table.Rows.Add(SplitTableRow(lines[i]));
        } else {
            for (int i = start; i <= end; i++) table.Rows.Add(SplitTableRow(lines[i]));
        }
        return table;
    }

    private static bool IsAlignmentRow(string line) {
        var cells = SplitTableRow(line);
        if (cells.Count == 0) return false;
        foreach (var c in cells) {
            var t = c.Trim(); if (t.Length < 3) return false;
            int dash = 0;
            for (int i = 0; i < t.Length; i++) {
                char ch = t[i];
                if (ch == '-') dash++;
                else if (ch == ':' && (i == 0 || i == t.Length - 1)) { } else return false;
            }
            if (dash < 3) return false;
        }
        return true;
    }

    private static ColumnAlignment ParseAlignmentCell(string cell) {
        var t = cell.Trim();
        if (t.StartsWith(":")) { if (t.EndsWith(":")) return ColumnAlignment.Center; return ColumnAlignment.Left; }
        if (t.EndsWith(":")) return ColumnAlignment.Right;
        return ColumnAlignment.None;
    }

    private static List<string> SplitTableRow(string line) {
        if (line is null) return new List<string>();
        var t = line.Trim();
        if (t.StartsWith("|")) t = t.Substring(1);
        if (t.EndsWith("|")) t = t.Substring(0, t.Length - 1);
        var parts = t.Split('|');
        var cells = new List<string>(parts.Length);
        foreach (var p in parts) cells.Add(p.Trim());
        return cells;
    }

    private static bool IsDefinitionLine(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var idx = line.IndexOf(':');
        if (idx <= 0) return false;
        if (idx + 1 >= line.Length) return false;
        return line[idx + 1] == ' ';
    }

    private static bool IsOrderedListLine(string line, out int number, out string content) {
        number = 0; content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        // Allow indentation; compute after leading spaces
        int spaces = 0; while (spaces < line.Length && line[spaces] == ' ') spaces++;
        int i = spaces; while (i < line.Length && char.IsDigit(line[i])) i++;
        if (i == spaces) return false;
        if (i < line.Length && line[i] == '.' && i + 1 < line.Length && line[i + 1] == ' ') {
            if (!int.TryParse(line.Substring(spaces, i - spaces), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 1;
            content = line.Substring(i + 2);
            return true;
        }
        return false;
    }

    private static bool IsOrderedListLine(string line, out int level, out int number, out string content) {
        level = 0; number = 0; content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        int spaces = 0; while (spaces < line.Length && line[spaces] == ' ') spaces++;
        int i = spaces; while (i < line.Length && char.IsDigit(line[i])) i++;
        if (i == spaces) return false;
        if (i < line.Length && line[i] == '.' && i + 1 < line.Length && line[i + 1] == ' ') {
            if (!int.TryParse(line.Substring(spaces, i - spaces), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 1;
            content = line.Substring(i + 2);
            level = spaces / 2;
            return true;
        }
        return false;
    }

    private static bool IsUnorderedListLine(string line, out bool isTask, out bool done, out string content) {
        isTask = false; done = false; content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.TrimStart();
        if (t.StartsWith("- ") || t.StartsWith("* ") || t.StartsWith("+ ")) {
            var c = t.Substring(2);
            if (c.StartsWith("[ ]")) { isTask = true; done = false; content = c.Length > 3 && c[2] == ']' && c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c; return true; }
            if (c.StartsWith("[x]", StringComparison.OrdinalIgnoreCase)) { isTask = true; done = true; content = c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c; return true; }
            content = c; return true;
        }
        return false;
    }

    private static bool IsUnorderedListLine(string line, out int level, out bool isTask, out bool done, out string content) {
        level = 0; isTask = false; done = false; content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        int spaces = 0; while (spaces < line.Length && line[spaces] == ' ') spaces++;
        if (spaces >= line.Length) return false;
        char ch = line[spaces];
        if ((ch == '-' || ch == '*' || ch == '+') && spaces + 1 < line.Length && line[spaces + 1] == ' ') {
            string c = line.Substring(spaces + 2);
            if (c.StartsWith("[ ]")) { isTask = true; done = false; content = c.Length > 3 && c[2] == ']' && c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c; level = spaces / 2; return true; }
            if (c.StartsWith("[x]", StringComparison.OrdinalIgnoreCase)) { isTask = true; done = true; content = c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c; level = spaces / 2; return true; }
            content = c; level = spaces / 2; return true;
        }
        return false;
    }

    private static bool IsCalloutHeader(string line, out string kind, out string title) {
        kind = string.Empty; title = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.TrimStart();
        if (!t.StartsWith(">")) return false;
        t = t.Substring(1).TrimStart();
        if (!t.StartsWith("[!")) return false;
        int close = t.IndexOf(']');
        if (close < 0 || close < 3) return false;
        string marker = t.Substring(2, close - 2);
        for (int i = 0; i < marker.Length; i++) if (!char.IsLetter(marker[i])) return false;
        kind = marker.ToLowerInvariant();
        title = t.Substring(close + 1).TrimStart();
        return title.Length > 0;
    }

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
