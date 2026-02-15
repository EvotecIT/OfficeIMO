using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight structured representation for a single page:
/// - Lines: plain text lines in top-to-bottom order
/// - Toc: table-of-contents style rows detected via dotted leaders
/// - ListItems: bullets and numbered list items
/// - LeaderRows: generic dotted-leader rows (label + trailing number)
/// - LinesDetailed: line geometry useful for higher-level extraction/debugging
/// - Tables: simple rows detected via large X gaps (heuristic)
/// </summary>
public sealed class StructuredPage {
    /// <summary>Plain text lines in natural reading order.</summary>
    public List<string> Lines { get; } = new();
    /// <summary>TOC entries: title + page number.</summary>
    public List<(string Title, int Page)> Toc { get; } = new();
    /// <summary>Bullet/numbered list items.</summary>
    public List<string> ListItems { get; } = new();
    /// <summary>Leader rows split into label and trailing number.</summary>
    public List<string[]> LeaderRows { get; } = new();
    /// <summary>Detected list nodes with hierarchical level.</summary>
    public List<StructuredListItem> ListNodes { get; } = new();
    /// <summary>Per-line geometry details (Y, XStart, XEnd, Text, Spans).</summary>
    public List<StructuredLine> LinesDetailed { get; } = new();
    /// <summary>Simple table-like rows derived from large X gaps per line.</summary>
    public List<string[]> Tables { get; } = new();
    /// <summary>Optional horizontal bands (line groups) for diagnostics/structure.</summary>
    public List<StructuredBand> Bands { get; } = new();
    /// <summary>Detailed tables with column geometry and band extents.</summary>
    public List<StructuredTable> TablesDetailed { get; } = new();
}

/// <summary>Represents a horizontal band grouping multiple lines.</summary>
public sealed class StructuredBand {
    /// <summary>Top Y (points) of the band (higher value is nearer top of page).</summary>
    public double YTop { get; init; }
    /// <summary>Bottom Y (points) of the band.</summary>
    public double YBottom { get; init; }
    /// <summary>Texts of lines grouped into this band in their original order.</summary>
    public List<string> Lines { get; init; } = new();
}

/// <summary>Represents a parsed list item (bullet or numbered) with hierarchy.</summary>
public sealed class StructuredListItem {
    /// <summary>1-based nesting level (best effort).</summary>
    public int Level { get; init; }
    /// <summary>Original marker like "1.2.3", "-", "•", "(a)".</summary>
    public string Marker { get; init; } = string.Empty;
    /// <summary>Normalized text of the list item.</summary>
    public string Text { get; init; } = string.Empty;
}

/// <summary>Table model with column geometry and extracted rows.</summary>
public sealed class StructuredTable {
    /// <summary>Top Y (points) of the band that produced this table.</summary>
    public double YTop { get; init; }
    /// <summary>Bottom Y (points) of the band that produced this table.</summary>
    public double YBottom { get; init; }
    /// <summary>Reason/heuristic for detection (e.g., band-splits, leaders).</summary>
    public string Kind { get; init; } = "band-splits";
    /// <summary>Detected columns with X ranges.</summary>
    public List<StructuredTableColumn> Columns { get; } = new();
    /// <summary>Extracted row values aligned to Columns.</summary>
    public List<string[]> Rows { get; } = new();
}

/// <summary>Column geometry for a detected table.</summary>
public sealed class StructuredTableColumn {
    /// <summary>Left X coordinate (points).</summary>
    public double From { get; init; }
    /// <summary>Right X coordinate (points).</summary>
    public double To { get; init; }
}
/// <summary>Geometry detail for a single emitted line.</summary>
public sealed class StructuredLine {
    /// <summary>Baseline Y coordinate for the line (points from bottom).</summary>
    public double Y { get; init; }
    /// <summary>Leftmost X coordinate (points).</summary>
    public double XStart { get; init; }
    /// <summary>Rightmost X coordinate (points).</summary>
    public double XEnd { get; init; }
    /// <summary>Line text.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Number of underlying spans grouped into this line.</summary>
    public int SpanCount { get; init; }
}

internal static class ContentStructureExtractor {
    private static readonly Regex TocRegex = new Regex(@"^(?<label>.+?)\s*\.{3,}\s+(?<num>\d{1,5})\s*$", RegexOptions.Compiled);
    private static readonly Regex ListRegex = new Regex(@"^\s*(?:[\u2022\-\*\u25CF]|\d+(?:\.\d+)*[\.)]|\([A-Za-z0-9]+\))\s+", RegexOptions.Compiled);
    private static readonly Regex NumberListRegex = new Regex(@"^\s*(?<mark>\d+(?:\.\d+)+)[\.)]?\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex BulletRegex = new Regex(@"^\s*(?<mark>[\u2022\-\*\u25CF])\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex ParenRegex = new Regex(@"^\s*\((?<mark>[A-Za-z0-9]+)\)\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly HashSet<string> CommonSuffixes = new(StringComparer.OrdinalIgnoreCase) {
        "ion", "ions", "ing", "ment", "tion", "sion", "iation", "ization",
        "ability", "ality", "able", "ible", "ance", "ence", "al", "ally",
        "er", "ers", "ed", "ly", "ology", "ologies"
    };

    public static StructuredPage Extract(IReadOnlyList<PdfTextSpan> spans, TextLayoutEngine.Options opts) {
        var page = new StructuredPage();
        var lines = TextLayoutEngine.BuildLines(spans, opts);
        var nonEmpty = new List<TextLayoutEngine.TextLine>();
        foreach (var ln in lines) if (!string.IsNullOrWhiteSpace(ln.Text)) nonEmpty.Add(ln);
        var bands = TextLayoutEngine.BandLines(nonEmpty, opts);
        // Fill detailed geometry first
        foreach (var ln in lines) {
            page.LinesDetailed.Add(new StructuredLine {
                Y = ln.Y,
                XStart = ln.XStart,
                XEnd = ln.XEnd,
                Text = ln.Text,
                SpanCount = ln.Spans.Count
            });
        }
        // Then semantic classification
        foreach (var ln in lines) {
            string t = ln.Text.Trim();
            if (t.Length == 0) continue;
            page.Lines.Add(t);
            var mToc = TocRegex.Match(t);
            if (mToc.Success && int.TryParse(mToc.Groups["num"].Value, out int num)) {
                var label = NormalizeShattered(mToc.Groups["label"].Value.TrimEnd('.').Trim());
                page.Toc.Add((label, num));
                page.LeaderRows.Add(new [] { label, num.ToString(System.Globalization.CultureInfo.InvariantCulture) });
                continue;
            }
            if (ListRegex.IsMatch(t)) {
                page.ListItems.Add(t);
                var mNum = NumberListRegex.Match(t);
                if (mNum.Success) {
                    string mark = mNum.Groups["mark"].Value;
                    int level = Math.Max(1, mark.Count(c => c == '.') + 1);
                    page.ListNodes.Add(new StructuredListItem { Level = level, Marker = mark, Text = mNum.Groups["text"].Value.Trim() });
                } else {
                    var mBul = BulletRegex.Match(t);
                    if (mBul.Success) page.ListNodes.Add(new StructuredListItem { Level = 1, Marker = mBul.Groups["mark"].Value, Text = mBul.Groups["text"].Value.Trim() });
                    else {
                        var mPar = ParenRegex.Match(t);
                        if (mPar.Success) page.ListNodes.Add(new StructuredListItem { Level = 1, Marker = "(" + mPar.Groups["mark"].Value + ")", Text = mPar.Groups["text"].Value.Trim() });
                    }
                }
            }
            else {
                // generic leader split: last whitespace + digits, after 3+ dots
                int dots = t.LastIndexOf("...", System.StringComparison.Ordinal);
                if (dots > 0) {
                    int k = t.Length - 1; while (k >= 0 && char.IsDigit(t[k])) k--; int numStart = k + 1;
                    if (numStart > 0 && numStart < t.Length && TryParsePositiveIntTail(t, numStart, out int n2)) {
                        var left = NormalizeShattered(t.Substring(0, numStart).TrimEnd('.', ' ').Trim());
                        var right = t.Substring(numStart);
                        page.LeaderRows.Add(new [] { left, right });
                    }
                }
            }
        }
        // Populate bands (diagnostics)
        foreach (var b in bands) {
            if (b.Count == 0) continue;
            double top = b[0].Y; double bottom = b[b.Count - 1].Y;
            var sb = new StructuredBand { YTop = top, YBottom = bottom };
            foreach (var ln in b) sb.Lines.Add(ln.Text);
            page.Bands.Add(sb);
        }

        // Table detection: prefer banded column inference; fallback to per-line
        var tables = TableDetector.DetectTablesFromBands(bands);
        if (tables.Count > 0) {
            // Clean leaders and add
            foreach (var t in tables) {
                if (string.Equals(t.Kind, "leaders", StringComparison.OrdinalIgnoreCase)) {
                    for (int r = 0; r < t.Rows.Count; r++) if (t.Rows[r].Length >= 2) {
                        t.Rows[r][0] = NormalizeShattered(t.Rows[r][0]);
                        t.Rows[r][1] = t.Rows[r][1].Trim('.');
                    }
                    // add only to detailed + LeaderRows; do NOT mix into generic Tables
                    page.TablesDetailed.Add(t);
                    foreach (var r in t.Rows) page.LeaderRows.Add(new[] { r[0], r[1] });
                    continue;
                }
                // Clean generic band/group tables to remove micro-token shattering and dot runs
                for (int r = 0; r < t.Rows.Count; r++) {
                    var row = t.Rows[r];
                    for (int c = 0; c < row.Length; c++) {
                        string cell = NormalizeShattered(row[c]);
                        // Normalize spacing around dots and collapse leader dot runs
                        cell = System.Text.RegularExpressions.Regex.Replace(cell, "\\s*\\.\\s*", ".");
                        int dotCount = 0; for (int k = 0; k < cell.Length; k++) if (cell[k] == '.') dotCount++;
                        if (dotCount >= 3) {
                            // Likely a leader run in this cell, drop the dots entirely
                            cell = cell.Replace(".", string.Empty).Trim();
                        }
                        t.Rows[r][c] = cell;
                    }
                }
                page.TablesDetailed.Add(t);
                page.Tables.AddRange(t.Rows);
            }
        } else {
            // Try a page-level leader-based table (TOC-like)
            var leaderTbl = TableDetector.DetectLeaderTable(nonEmpty);
            if (leaderTbl is not null) {
                if (string.Equals(leaderTbl.Kind, "leaders", StringComparison.OrdinalIgnoreCase)) {
                    for (int r = 0; r < leaderTbl.Rows.Count; r++) if (leaderTbl.Rows[r].Length >= 2) {
                        leaderTbl.Rows[r][0] = NormalizeShattered(leaderTbl.Rows[r][0]);
                        leaderTbl.Rows[r][1] = leaderTbl.Rows[r][1].Trim('.');
                    }
                }
                page.TablesDetailed.Add(leaderTbl);
                foreach (var r in leaderTbl.Rows) page.LeaderRows.Add(new [] { r[0], r[1] });
            } else {
                var rows = TableDetector.Detect(lines);
                if (rows.Count > 0) {
                    foreach (var r in rows) if (r.Length >= 2) { r[0] = NormalizeShattered(r[0]); r[1] = r[1].Trim('.'); }
                    page.Tables.AddRange(rows);
                }
            }
        }
        return page;
    }

    private static bool IsWordish(char c) => char.IsLetter(c) || c == '\'' || c == '-' || c == '/';
    private static bool IsAllLetters(string s) { for (int i = 0; i < s.Length; i++) if (!IsWordish(s[i])) return false; return s.Length > 0; }
    private static bool IsShortAbbrev(string s) {
        if (s.Length == 0 || s.Length > 3) return false;
        for (int i = 0; i < s.Length; i++) if (!char.IsUpper(s[i])) return false; return true;
    }
    private static string NormalizeShattered(string s) {
        if (string.IsNullOrEmpty(s)) return s;
        // Collapse runs of spaces
        s = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
        var parts = s.Split(' ');
        if (parts.Length <= 2) {
            if (parts.Length == 2 && IsAllLetters(parts[0]) && IsAllLetters(parts[1])) {
                if (parts[0].Length == 1 && parts[1].Length >= 3) return parts[0] + parts[1];
                if (parts[1].Length <= 2 || parts[0].Length <= 2) return parts[0] + parts[1];
            }
            return s;
        }
        int shortCount = parts.Count(p => p.Length <= 2 && IsAllLetters(p));
        // Heuristic: only join when there are multiple micro tokens
        if (shortCount < 2) return s; // mostly healthy
        var sb = new System.Text.StringBuilder(s.Length);
        sb.Append(parts[0]);
        for (int i = 1; i < parts.Length; i++) {
            string prev = parts[i - 1]; string cur = parts[i];
            bool upperSinglesJoin = prev.Length == 1 && cur.Length == 1 && char.IsUpper(prev[0]) && char.IsUpper(cur[0])
                                    && (i + 1 < parts.Length && parts[i + 1].Length == 1 && char.IsUpper(parts[i + 1][0]));
            bool leadingLetterJoin = IsAllLetters(prev) && IsAllLetters(cur) && prev.Length == 1 && cur.Length >= 3;
            bool lettersJoin = IsAllLetters(prev) && IsAllLetters(cur) && ((prev.Length <= 2 || cur.Length <= 2) || leadingLetterJoin || upperSinglesJoin) && !(IsShortAbbrev(prev) && IsShortAbbrev(cur) && !upperSinglesJoin);
            // lookahead to aggressively join runs of short tokens
            bool nextShort = (i + 1 < parts.Length) && parts[i + 1].Length <= 2 && IsAllLetters(parts[i + 1]) && !IsShortAbbrev(parts[i + 1]);
            if (lettersJoin || (IsAllLetters(cur) && cur.Length <= 2 && nextShort)) sb.Append(cur);
            else sb.Append(' ').Append(cur);
        }
        string joined = sb.ToString().Replace("  ", " ");
        // Secondary pass: join common suffix fragments (e.g., "Except ions" -> "Exceptions")
        var toks = joined.Split(' ');
        if (toks.Length > 1) {
            var sb2 = new System.Text.StringBuilder(joined.Length);
            sb2.Append(toks[0]);
            for (int i = 1; i < toks.Length; i++) {
                string prev = toks[i - 1]; string cur = toks[i];
                if (IsAllLetters(prev) && IsAllLetters(cur) && CommonSuffixes.Contains(cur)) sb2.Append(cur);
                else sb2.Append(' ').Append(cur);
            }
            joined = sb2.ToString();
        }
        return joined;
    }

    private static bool TryParsePositiveIntTail(string text, int startIndex, out int value) {
        value = 0;
        if (string.IsNullOrEmpty(text)) return false;
        if (startIndex < 0 || startIndex >= text.Length) return false;

        for (int i = startIndex; i < text.Length; i++) {
            var c = text[i];
            if (c < '0' || c > '9') return false;
            var digit = c - '0';
            if (value > ((int.MaxValue - digit) / 10)) return false;
            value = (value * 10) + digit;
        }
        return true;
    }
}
