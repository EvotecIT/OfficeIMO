using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight structured representation for a single page:
/// - Lines: plain text lines in top-to-bottom order
/// - Toc: table-of-contents style rows detected via dotted leaders
/// - ListItems: bullets and numbered list items
/// - LeaderRows: generic leader rows (label + trailing value)
/// - LinesDetailed: line geometry useful for higher-level extraction/debugging
/// - Headings: heuristic heading lines inferred from larger-than-body font sizes
/// - Paragraphs: heuristic paragraph groups built from nearby non-list, non-table lines
/// - Tables: simple rows detected via large X gaps (heuristic)
/// </summary>
public sealed class StructuredPage {
    private readonly HashSet<(string Label, string Value)> _leaderRowKeys = new();
    /// <summary>Plain text lines in natural reading order.</summary>
    public List<string> Lines { get; } = new();
    /// <summary>TOC entries: title + page number.</summary>
    public List<(string Title, int Page)> Toc { get; } = new();
    /// <summary>Bullet/numbered list items.</summary>
    public List<string> ListItems { get; } = new();
    /// <summary>Leader rows split into label and trailing value.</summary>
    public List<string[]> LeaderRows { get; } = new();

    internal bool TryAddLeaderRow(string label, string value) {
        if (!_leaderRowKeys.Add((label, value))) {
            return false;
        }

        LeaderRows.Add(new[] { label, value });
        return true;
    }
    /// <summary>Detected list nodes with hierarchical level.</summary>
    public List<StructuredListItem> ListNodes { get; } = new();
    /// <summary>Per-line geometry details (Y, XStart, XEnd, Text, Spans).</summary>
    public List<StructuredLine> LinesDetailed { get; } = new();
    /// <summary>Heuristic heading lines inferred from larger-than-body font sizes.</summary>
    public List<StructuredHeading> Headings { get; } = new();
    /// <summary>Heuristic paragraph groups built from nearby non-list, non-table lines.</summary>
    public List<StructuredParagraph> Paragraphs { get; } = new();
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
    public double YTop { get; set; }
    /// <summary>Bottom Y (points) of the band.</summary>
    public double YBottom { get; set; }
    /// <summary>Texts of lines grouped into this band in their original order.</summary>
    public List<string> Lines { get; set; } = new();
}

/// <summary>Represents a parsed list item (bullet or numbered) with hierarchy.</summary>
public sealed class StructuredListItem {
    /// <summary>1-based nesting level (best effort).</summary>
    public int Level { get; set; }
    /// <summary>Original marker like "1.2.3", "-", "•", "(a)".</summary>
    public string Marker { get; set; } = string.Empty;
    /// <summary>Normalized text of the list item.</summary>
    public string Text { get; set; } = string.Empty;
    /// <summary>Line geometry for the source list item.</summary>
    public StructuredLine Line { get; set; } = new();
}

/// <summary>Table model with column geometry and extracted rows.</summary>
public sealed class StructuredTable {
    /// <summary>Top Y (points) of the band that produced this table.</summary>
    public double YTop { get; set; }
    /// <summary>Bottom Y (points) of the band that produced this table.</summary>
    public double YBottom { get; set; }
    /// <summary>Reason/heuristic for detection (e.g., band-splits, leaders).</summary>
    public string Kind { get; set; } = "band-splits";
    /// <summary>Detected columns with X ranges.</summary>
    public List<StructuredTableColumn> Columns { get; } = new();
    /// <summary>Extracted row values aligned to Columns.</summary>
    public List<string[]> Rows { get; } = new();
}

/// <summary>Column geometry for a detected table.</summary>
public sealed class StructuredTableColumn {
    /// <summary>Left X coordinate (points).</summary>
    public double From { get; set; }
    /// <summary>Right X coordinate (points).</summary>
    public double To { get; set; }
}

/// <summary>Detected tables for a single document page.</summary>
public sealed class StructuredTablePage {
    /// <summary>Creates a page table result.</summary>
    public StructuredTablePage(int pageNumber, IEnumerable<StructuredTable> tables) {
        if (pageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        PageNumber = pageNumber;
        Tables.AddRange(tables ?? throw new ArgumentNullException(nameof(tables)));
    }

    /// <summary>1-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Detected tables on this page.</summary>
    public List<StructuredTable> Tables { get; } = new();
}

/// <summary>Detected paragraphs for a single document page.</summary>
public sealed class StructuredParagraphPage {
    /// <summary>Creates a page paragraph result.</summary>
    public StructuredParagraphPage(int pageNumber, IEnumerable<StructuredParagraph> paragraphs) {
        if (pageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        PageNumber = pageNumber;
        Paragraphs.AddRange(paragraphs ?? throw new ArgumentNullException(nameof(paragraphs)));
    }

    /// <summary>1-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Detected paragraphs on this page.</summary>
    public List<StructuredParagraph> Paragraphs { get; } = new();
}

/// <summary>Detected headings for a single document page.</summary>
public sealed class StructuredHeadingPage {
    /// <summary>Creates a page heading result.</summary>
    public StructuredHeadingPage(int pageNumber, IEnumerable<StructuredHeading> headings) {
        if (pageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        PageNumber = pageNumber;
        Headings.AddRange(headings ?? throw new ArgumentNullException(nameof(headings)));
    }

    /// <summary>1-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Detected headings on this page.</summary>
    public List<StructuredHeading> Headings { get; } = new();
}

/// <summary>Detected list items for a single document page.</summary>
public sealed class StructuredListItemPage {
    /// <summary>Creates a page list-item result.</summary>
    public StructuredListItemPage(int pageNumber, IEnumerable<StructuredListItem> listItems) {
        if (pageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        PageNumber = pageNumber;
        ListItems.AddRange(listItems ?? throw new ArgumentNullException(nameof(listItems)));
    }

    /// <summary>1-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Detected list items on this page.</summary>
    public List<StructuredListItem> ListItems { get; } = new();
}

/// <summary>Heuristic heading line inferred from font size and geometry.</summary>
public sealed class StructuredHeading {
    /// <summary>Best-effort heading level, where 1 is the largest heading tier.</summary>
    public int Level { get; set; }
    /// <summary>Heading text.</summary>
    public string Text { get; set; } = string.Empty;
    /// <summary>Line geometry for the heading.</summary>
    public StructuredLine Line { get; set; } = new();
    /// <summary>Representative font size in points.</summary>
    public double FontSize { get; set; }
}

/// <summary>Heuristic paragraph group built from nearby non-list, non-table lines.</summary>
public sealed class StructuredParagraph {
    /// <summary>Paragraph text with grouped lines joined by spaces.</summary>
    public string Text { get; set; } = string.Empty;
    /// <summary>Line geometry entries that make up the paragraph.</summary>
    public List<StructuredLine> Lines { get; } = new();
    /// <summary>Leftmost X coordinate (points).</summary>
    public double XStart { get; set; }
    /// <summary>Rightmost X coordinate (points).</summary>
    public double XEnd { get; set; }
    /// <summary>Top baseline Y coordinate (points).</summary>
    public double YTop { get; set; }
    /// <summary>Bottom baseline Y coordinate (points).</summary>
    public double YBottom { get; set; }
}

/// <summary>Geometry detail for a single emitted line.</summary>
public sealed class StructuredLine {
    /// <summary>Baseline Y coordinate for the line (points from bottom).</summary>
    public double Y { get; set; }
    /// <summary>Leftmost X coordinate (points).</summary>
    public double XStart { get; set; }
    /// <summary>Rightmost X coordinate (points).</summary>
    public double XEnd { get; set; }
    /// <summary>Line text.</summary>
    public string Text { get; set; } = string.Empty;
    /// <summary>Representative font size in points.</summary>
    public double FontSize { get; set; }
    /// <summary>Number of underlying spans grouped into this line.</summary>
    public int SpanCount { get; set; }
}

internal static class ContentStructureExtractor {
    private static readonly Regex ListRegex = new Regex(@"^\s*(?:[\u2022\u25CF]\s*|[\-\*]\s+|\d+(?:\.\d+)*[\.)]\s*|\([A-Za-z0-9]+\)\s*).+", RegexOptions.Compiled);
    private static readonly Regex NumberListRegex = new Regex(@"^\s*(?<mark>\d+(?:\.\d+)*)[\.)]\s*(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex BulletRegex = new Regex(@"^\s*(?:(?<mark>[\u2022\u25CF])\s*|(?<mark>[\-\*])\s+)(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex ParenRegex = new Regex(@"^\s*\((?<mark>[A-Za-z0-9]+)\)\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly HashSet<string> CommonSuffixes = new(StringComparer.OrdinalIgnoreCase) {
        "ion", "ions", "ing", "ment", "tion", "sion", "iation", "ization",
        "ability", "ality", "able", "ible", "ance", "ence", "al", "ally",
        "er", "ers", "ed", "ly", "ology", "ologies"
    };

    public static StructuredPage Extract(IReadOnlyList<PdfTextSpan> spans, TextLayoutEngine.Options opts) {
        var page = new StructuredPage();
        var fallbackTableLines = new HashSet<TextLayoutEngine.TextLine>();
        var lines = TextLayoutEngine.BuildLines(spans, opts);
        var nonEmpty = new List<TextLayoutEngine.TextLine>();
        foreach (var ln in lines) if (!string.IsNullOrWhiteSpace(ln.Text)) nonEmpty.Add(ln);
        var bands = TextLayoutEngine.BandLines(nonEmpty, opts);
        // Fill detailed geometry first
        foreach (var ln in lines) {
            page.LinesDetailed.Add(ToStructuredLine(ln));
        }
        // Then semantic classification
        foreach (var ln in lines) {
            string t = ln.Text.Trim();
            if (t.Length == 0) continue;
            page.Lines.Add(t);
            if (TryParseTocRow(t, out string tocLabel, out int num)) {
                var label = NormalizeShattered(tocLabel.TrimEnd('.').Trim());
                page.Toc.Add((label, num));
                AddLeaderRow(page, label, num.ToString(System.Globalization.CultureInfo.InvariantCulture));
                continue;
            }
            if (ListRegex.IsMatch(t)) {
                page.ListItems.Add(t);
                var mNum = NumberListRegex.Match(t);
                if (mNum.Success) {
                    string mark = mNum.Groups["mark"].Value;
                    int level = Math.Max(1, mark.Count(c => c == '.') + 1);
                    page.ListNodes.Add(new StructuredListItem { Level = level, Marker = mark, Text = mNum.Groups["text"].Value.Trim(), Line = ToStructuredLine(ln) });
                } else {
                    var mBul = BulletRegex.Match(t);
                    if (mBul.Success) page.ListNodes.Add(new StructuredListItem { Level = 1, Marker = mBul.Groups["mark"].Value, Text = mBul.Groups["text"].Value.Trim(), Line = ToStructuredLine(ln) });
                    else {
                        var mPar = ParenRegex.Match(t);
                        if (mPar.Success) page.ListNodes.Add(new StructuredListItem { Level = 1, Marker = "(" + mPar.Groups["mark"].Value + ")", Text = mPar.Groups["text"].Value.Trim(), Line = ToStructuredLine(ln) });
                    }
                }
            }
            else {
                if (TryParseLeaderRow(t, out string leaderLabel, out string leaderValue)) {
                    var value = NormalizeLeaderValue(leaderValue);
                    if (value.Length > 0) {
                        var left = NormalizeShattered(leaderLabel.TrimEnd('.', '-', '_', ' ').Trim());
                        AddLeaderRow(page, left, value);
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
                        t.Rows[r][1] = NormalizeLeaderValue(t.Rows[r][1]);
                    }
                    // add only to detailed + LeaderRows; do NOT mix into generic Tables
                    page.TablesDetailed.Add(t);
                    foreach (var r in t.Rows) AddLeaderRow(page, r[0], r[1]);
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
                        leaderTbl.Rows[r][1] = NormalizeLeaderValue(leaderTbl.Rows[r][1]);
                    }
                }
                page.TablesDetailed.Add(leaderTbl);
                foreach (var r in leaderTbl.Rows) AddLeaderRow(page, r[0], r[1]);
            } else {
                var rows = TableDetector.DetectLineRows(lines);
                if (rows.Count > 0) {
                    foreach (var row in rows) {
                        var r = row.Cells;
                        if (r.Length >= 2) {
                            r[0] = NormalizeShattered(r[0]);
                            r[1] = r[1].Trim('.');
                        }

                        fallbackTableLines.Add(row.Line);
                        page.Tables.Add(r);
                    }
                }
            }
        }
        AddHeadings(page, nonEmpty);
        AddParagraphs(page, nonEmpty, fallbackTableLines);
        return page;
    }

    private static void AddParagraphs(StructuredPage page, List<TextLayoutEngine.TextLine> lines, HashSet<TextLayoutEngine.TextLine> fallbackTableLines) {
        var candidates = new List<TextLayoutEngine.TextLine>();
        foreach (var line in lines) {
            string text = line.Text.Trim();
            if (text.Length == 0 ||
                ListRegex.IsMatch(text) ||
                IsHeadingLine(line, page.Headings) ||
                fallbackTableLines.Contains(line) ||
                IsInsideTable(line, page.TablesDetailed)) {
                continue;
            }

            candidates.Add(line);
        }

        if (candidates.Count == 0) {
            return;
        }

        var gaps = new List<double>();
        for (int i = 1; i < candidates.Count; i++) {
            double gap = candidates[i - 1].Y - candidates[i].Y;
            if (gap > 0.001) {
                gaps.Add(gap);
            }
        }

        double medianGap = Median(gaps);
        double splitGap = medianGap <= 0 ? 18D : Math.Max(18D, medianGap * 1.35D);
        double xTolerance = 18D;
        var current = new List<TextLayoutEngine.TextLine> { candidates[0] };

        for (int i = 1; i < candidates.Count; i++) {
            var previous = candidates[i - 1];
            var next = candidates[i];
            double gap = previous.Y - next.Y;
            bool split = gap > splitGap || Math.Abs(next.XStart - current[0].XStart) > xTolerance;
            if (split) {
                page.Paragraphs.Add(BuildParagraph(current));
                current = new List<TextLayoutEngine.TextLine>();
            }

            current.Add(next);
        }

        if (current.Count > 0) {
            page.Paragraphs.Add(BuildParagraph(current));
        }
    }

    private static StructuredParagraph BuildParagraph(List<TextLayoutEngine.TextLine> lines) {
        var paragraph = new StructuredParagraph {
            Text = string.Join(" ", lines.Select(line => line.Text.Trim())),
            XStart = lines.Min(line => line.XStart),
            XEnd = lines.Max(line => line.XEnd),
            YTop = lines.Max(line => line.Y),
            YBottom = lines.Min(line => line.Y)
        };

        for (int i = 0; i < lines.Count; i++) {
            var line = lines[i];
            paragraph.Lines.Add(new StructuredLine {
                Y = line.Y,
                XStart = line.XStart,
                XEnd = line.XEnd,
                Text = line.Text,
                FontSize = GetLineFontSize(line),
                SpanCount = line.Spans.Count
            });
        }

        return paragraph;
    }

    private static void AddHeadings(StructuredPage page, List<TextLayoutEngine.TextLine> lines) {
        var bodySizes = new List<double>();
        foreach (var line in lines) {
            string text = line.Text.Trim();
            if (text.Length == 0 ||
                ListRegex.IsMatch(text) ||
                IsInsideTable(line, page.TablesDetailed)) {
                continue;
            }

            bodySizes.Add(GetLineFontSize(line));
        }

        double bodySize = EstimateBodyFontSize(bodySizes);
        if (bodySize <= 0) {
            return;
        }

        foreach (var line in lines) {
            string text = line.Text.Trim();
            if (text.Length == 0 ||
                text.Length > 160 ||
                ListRegex.IsMatch(text) ||
                IsInsideTable(line, page.TablesDetailed)) {
                continue;
            }

            double fontSize = GetLineFontSize(line);
            if (fontSize < Math.Max(bodySize + 1.5D, bodySize * 1.18D)) {
                continue;
            }

            var structuredLine = new StructuredLine {
                Y = line.Y,
                XStart = line.XStart,
                XEnd = line.XEnd,
                Text = text,
                FontSize = fontSize,
                SpanCount = line.Spans.Count
            };
            page.Headings.Add(new StructuredHeading {
                Level = GetHeadingLevel(fontSize, bodySize),
                Text = text,
                Line = structuredLine,
                FontSize = fontSize
            });
        }
    }

    private static int GetHeadingLevel(double fontSize, double bodySize) {
        if (fontSize >= bodySize * 1.65D) {
            return 1;
        }

        if (fontSize >= bodySize * 1.35D) {
            return 2;
        }

        return 3;
    }

    private static double EstimateBodyFontSize(List<double> fontSizes) {
        if (fontSizes.Count == 0) {
            return 0D;
        }

        fontSizes.Sort();
        int index = Math.Max(0, (int)Math.Floor((fontSizes.Count - 1) * 0.35D));
        return fontSizes[index];
    }

    private static bool IsHeadingLine(TextLayoutEngine.TextLine line, List<StructuredHeading> headings) {
        for (int i = 0; i < headings.Count; i++) {
            var heading = headings[i];
            if (Math.Abs(heading.Line.Y - line.Y) <= 0.001 &&
                Math.Abs(heading.Line.XStart - line.XStart) <= 0.001 &&
                string.Equals(heading.Text, line.Text.Trim(), StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static double GetLineFontSize(TextLayoutEngine.TextLine line) {
        double fontSize = 0D;
        for (int i = 0; i < line.Spans.Count; i++) {
            fontSize = Math.Max(fontSize, line.Spans[i].FontSize);
        }

        return fontSize;
    }

    private static StructuredLine ToStructuredLine(TextLayoutEngine.TextLine line) {
        return new StructuredLine {
            Y = line.Y,
            XStart = line.XStart,
            XEnd = line.XEnd,
            Text = line.Text,
            FontSize = GetLineFontSize(line),
            SpanCount = line.Spans.Count
        };
    }

    private static bool IsInsideTable(TextLayoutEngine.TextLine line, List<StructuredTable> tables) {
        for (int i = 0; i < tables.Count; i++) {
            var table = tables[i];
            if (table.Columns.Count == 0) {
                continue;
            }

            if (line.Y <= table.YTop + 0.001 &&
                line.Y >= table.YBottom - 0.001 &&
                line.XEnd >= table.Columns[0].From - 2D) {
                return true;
            }
        }

        return false;
    }

    private static double Median(List<double> values) {
        if (values.Count == 0) {
            return 0D;
        }

        values.Sort();
        int middle = values.Count / 2;
        if ((values.Count & 1) == 1) {
            return values[middle];
        }

        return (values[middle - 1] + values[middle]) / 2D;
    }

    private static void AddLeaderRow(StructuredPage page, string label, string value) {
        label = NormalizeShattered(label ?? string.Empty).Trim();
        value = NormalizeLeaderValue(value);
        if (label.Length == 0 || value.Length == 0) {
            return;
        }

        page.TryAddLeaderRow(label, value);
    }

    private static bool TryParseTocRow(string text, out string label, out int pageNumber) {
        label = string.Empty;
        pageNumber = 0;
        int trailingContentEnd = text.Length;
        while (trailingContentEnd > 0 && char.IsWhiteSpace(text[trailingContentEnd - 1])) {
            trailingContentEnd--;
        }

        for (int index = 1; index < text.Length;) {
            if (text[index] != '.') {
                index++;
                continue;
            }

            int runStart = index;
            while (index < text.Length && text[index] == '.') {
                index++;
            }

            if (index - runStart < 3 || index >= text.Length || !char.IsWhiteSpace(text[index])) {
                continue;
            }

            int digitStart = SkipWhitespace(text, index);
            int digitEnd = digitStart;
            while (digitEnd < trailingContentEnd && digitEnd - digitStart < 6 && text[digitEnd] is >= '0' and <= '9') {
                digitEnd++;
            }

            int digitCount = digitEnd - digitStart;
            if (digitCount is < 1 or > 5 || digitEnd != trailingContentEnd) {
                continue;
            }

            label = text.Substring(0, runStart);
            for (int digitIndex = digitStart; digitIndex < digitEnd; digitIndex++) {
                pageNumber = checked((pageNumber * 10) + (text[digitIndex] - '0'));
            }

            return true;
        }

        return false;
    }

    private static bool TryParseLeaderRow(string text, out string label, out string value) {
        label = string.Empty;
        value = string.Empty;
        int validValueSuffixStart = FindValidLeaderValueSuffixStart(text);

        for (int index = 1; index < text.Length;) {
            char leader = text[index];
            if (leader != '.' && leader != '-' && leader != '_') {
                index++;
                continue;
            }

            int runStart = index;
            while (index < text.Length && text[index] == leader) {
                index++;
            }

            if (index - runStart < 3) {
                continue;
            }

            int valueStart = SkipWhitespace(text, index);
            if (valueStart < text.Length && IsLeaderCurrency(text[valueStart])) {
                valueStart = SkipWhitespace(text, valueStart + 1);
            }

            if (valueStart >= text.Length ||
                !char.IsLetterOrDigit(text[valueStart]) ||
                valueStart < validValueSuffixStart) {
                continue;
            }

            label = text.Substring(0, runStart);
            value = text.Substring(index);
            return true;
        }

        return false;
    }

    private static int FindValidLeaderValueSuffixStart(string text) {
        for (int index = text.Length - 1; index >= 0; index--) {
            char character = text[index];
            bool allowed = char.IsLetterOrDigit(character) ||
                           char.IsWhiteSpace(character) ||
                           character is '.' or ',' or '\'' or '/' or '%' or '+' or '-' or '(' or ')';
            if (!allowed) {
                return index + 1;
            }
        }

        return 0;
    }

    private static int SkipWhitespace(string text, int index) {
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }

        return index;
    }

    private static bool IsLeaderCurrency(char value) => value is '$' or '€' or '£';

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

    private static string NormalizeLeaderValue(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        string normalized = Regex.Replace(value!.Trim(), "\\s+", " ");
        normalized = Regex.Replace(normalized, "\\s*([.,])\\s*", "$1");
        normalized = Regex.Replace(normalized, "([$€£])\\s+", "$1");
        normalized = normalized.Trim('.');

        bool hasDigit = false;
        for (int i = 0; i < normalized.Length; i++) {
            if (char.IsDigit(normalized[i])) {
                hasDigit = true;
                break;
            }
        }

        return hasDigit ? normalized : string.Empty;
    }
}
