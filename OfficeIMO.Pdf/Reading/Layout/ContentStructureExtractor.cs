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
    internal IReadOnlyList<PdfTextSpan> SourceRuns { get; set; } = Array.Empty<PdfTextSpan>();
    internal IReadOnlyList<TextLayoutEngine.TextLine> SourceLines { get; set; } =
        Array.Empty<TextLayoutEngine.TextLine>();
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
    /// <summary>Normalized source-extraction confidence from 0 through 1.</summary>
    public double Confidence { get; internal set; } = 1D;
    /// <summary>Immutable snapshot of the positioned text spans grouped into this line.</summary>
    public IReadOnlyList<PdfTextSpan> Spans { get; internal set; } = Array.Empty<PdfTextSpan>();
    internal PdfLogicalContentSourceKind SourceKind { get; set; } = PdfLogicalContentSourceKind.Native;
    internal PdfLogicalVisualBounds? VisualBounds { get; set; }
    /// <summary>Number of underlying spans grouped into this line.</summary>
    public int SpanCount => Spans.Count;
}

internal static class ContentStructureExtractor {
    private readonly struct StructuredTableBounds {
        internal StructuredTableBounds(double top, double bottom, double left, double right) {
            Top = top;
            Bottom = bottom;
            Left = left;
            Right = right;
        }

        internal double Top { get; }
        internal double Bottom { get; }
        internal double Left { get; }
        internal double Right { get; }
    }

    private const RegexOptions StructuralRegexOptions = RegexOptions.Compiled | RegexOptions.CultureInvariant;
    private static readonly Regex BulletRegex = new Regex(@"^\s*(?:(?<mark>[\u2022\u2023\u2043\u2219\u25AA\u25AB\u25CF\u25CB\u25E6])\s*|(?<mark>[\-\*])(?:\s+|(?![\-\*])(?!(?:\p{Sc}\s*)?(?:\p{Nd}|[\.,]\p{Nd}))(?=[^\p{Nd}\s])))(?<text>.+)$", StructuralRegexOptions);

    internal static bool IsListItemText(string text) =>
        TryParseListItemText(text, out _, out _, out _);

    internal static bool IsSentenceTerminal(char value) => value is
        '.' or '!' or '?' or ':' or
        '\u0589' or // Armenian full stop
        '\u061F' or '\u06D4' or // Arabic question mark and full stop
        '\u0964' or '\u0965' or // danda and double danda
        '\u1362' or '\u1367' or '\u1368' or // Ethiopic sentence punctuation
        '\u166E' or '\u1803' or '\u1809' or
        '\u203C' or '\u2047' or '\u2048' or '\u2049' or '\u2E2E' or
        '\u3002' or '\uFE52' or '\uFE56' or '\uFE57' or
        '\uFF01' or '\uFF0E' or '\uFF1A' or '\uFF1F';

    internal static bool EndsWithSentenceTerminal(string text) {
        for (int index = text.Length - 1; index >= 0; index--) {
            char value = text[index];
            if (char.IsWhiteSpace(value)) continue;
            if (IsSentenceTerminal(value)) return true;

            System.Globalization.UnicodeCategory category = char.GetUnicodeCategory(value);
            if (category == System.Globalization.UnicodeCategory.ClosePunctuation ||
                category == System.Globalization.UnicodeCategory.FinalQuotePunctuation) continue;
            return false;
        }
        return false;
    }

    internal static bool TryParseListItemText(
        string text,
        out string marker,
        out string body,
        out int level) {
        marker = string.Empty;
        body = string.Empty;
        level = 1;
        if (string.IsNullOrWhiteSpace(text)) return false;

        if (TryParseNumberedListItem(text, out marker, out body, out level)) return true;

        Match bullet = BulletRegex.Match(text);
        if (bullet.Success) {
            marker = bullet.Groups["mark"].Value;
            body = bullet.Groups["text"].Value.Trim();
            return body.Length > 0;
        }

        if (TryParseParenthesizedListItem(text, out marker, out body)) return true;
        return TryParseIdeographicListItem(text, out marker, out body);
    }

    private static bool TryParseNumberedListItem(
        string text,
        out string marker,
        out string body,
        out int level) {
        marker = string.Empty;
        body = string.Empty;
        level = 1;
        int index = SkipWhitespace(text, 0);
        int markerStart = index;
        if (!TryConsumeDecimalDigits(text, ref index)) return false;

        int markerEnd = -1;
        while (index < text.Length && text[index] is '.' or '\uFF0E') {
            int separator = index++;
            if (TryConsumeDecimalDigits(text, ref index)) {
                level++;
                continue;
            }

            markerEnd = separator;
            break;
        }

        if (markerEnd < 0) {
            if (index >= text.Length || text[index] is not (')' or '\u3001' or '\uFF09')) return false;
            markerEnd = index++;
        }
        if (index < text.Length && TryGetDecimalDigit(text, index, out _, out _)) return false;
        marker = text.Substring(markerStart, markerEnd - markerStart);
        body = text.Substring(SkipWhitespace(text, index)).Trim();
        return body.Length > 0;
    }

    private static bool TryParseParenthesizedListItem(
        string text,
        out string marker,
        out string body) {
        marker = string.Empty;
        body = string.Empty;
        int index = SkipWhitespace(text, 0);
        if (index >= text.Length || text[index] is not ('(' or '\uFF08')) return false;
        int markerStart = index++;
        int scalarCount = 0;
        bool allDigits = true;
        while (index < text.Length && text[index] is not (')' or '\uFF09')) {
            System.Globalization.UnicodeCategory category =
                System.Globalization.CharUnicodeInfo.GetUnicodeCategory(text, index);
            bool isDigit = System.Globalization.CharUnicodeInfo.GetDecimalDigitValue(text, index) >= 0;
            bool isLetter = category is
                System.Globalization.UnicodeCategory.UppercaseLetter or
                System.Globalization.UnicodeCategory.LowercaseLetter or
                System.Globalization.UnicodeCategory.TitlecaseLetter or
                System.Globalization.UnicodeCategory.ModifierLetter or
                System.Globalization.UnicodeCategory.OtherLetter or
                System.Globalization.UnicodeCategory.LetterNumber;
            if (!isDigit && !isLetter) return false;
            allDigits &= isDigit;
            scalarCount++;
            index += char.IsSurrogatePair(text, index) ? 2 : 1;
            if (scalarCount > 4) return false;
        }
        if (index >= text.Length || scalarCount == 0 || (!allDigits && scalarCount != 1)) return false;
        index++;
        marker = text.Substring(markerStart, index - markerStart);
        body = text.Substring(SkipWhitespace(text, index)).Trim();
        return body.Length > 0;
    }

    private static bool TryParseIdeographicListItem(
        string text,
        out string marker,
        out string body) {
        marker = string.Empty;
        body = string.Empty;
        int index = SkipWhitespace(text, 0);
        if (index >= text.Length) return false;
        int markerStart = index;
        System.Globalization.UnicodeCategory category =
            System.Globalization.CharUnicodeInfo.GetUnicodeCategory(text, index);
        bool isMarker = System.Globalization.CharUnicodeInfo.GetDecimalDigitValue(text, index) >= 0 ||
            category is
                System.Globalization.UnicodeCategory.UppercaseLetter or
                System.Globalization.UnicodeCategory.LowercaseLetter or
                System.Globalization.UnicodeCategory.TitlecaseLetter or
                System.Globalization.UnicodeCategory.ModifierLetter or
                System.Globalization.UnicodeCategory.OtherLetter or
                System.Globalization.UnicodeCategory.LetterNumber;
        if (!isMarker) return false;
        index += char.IsSurrogatePair(text, index) ? 2 : 1;
        int markerEnd = index;
        if (index >= text.Length || text[index++] != '\u3001') return false;
        marker = text.Substring(markerStart, markerEnd - markerStart);
        body = text.Substring(SkipWhitespace(text, index)).Trim();
        return body.Length > 0;
    }

    private static bool TryConsumeDecimalDigits(string text, ref int index) {
        int start = index;
        while (index < text.Length && TryGetDecimalDigit(text, index, out _, out int consumed)) {
            index += consumed;
        }
        return index > start;
    }

    public static StructuredPage Extract(
        IReadOnlyList<PdfTextSpan> spans,
        TextLayoutEngine.Options opts,
        double? pageHeight = null) =>
        Extract(spans, opts, pageHeight, consumeWork: null, cancellationCheck: null);

    internal static StructuredPage Extract(
        IReadOnlyList<PdfTextSpan> spans,
        TextLayoutEngine.Options opts,
        double? pageHeight,
        Action<long>? consumeWork,
        Action? cancellationCheck,
        IReadOnlyList<StructuredTable>? precomputedTables = null) {
        cancellationCheck?.Invoke();
        var page = new StructuredPage();
        var fallbackTableLines = new HashSet<TextLayoutEngine.TextLine>();
        var lines = TextLayoutEngine.BuildLines(spans, opts, consumeWork, cancellationCheck);
        var nonEmpty = new List<TextLayoutEngine.TextLine>();
        foreach (var ln in lines) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            if (!string.IsNullOrWhiteSpace(ln.Text)) nonEmpty.Add(ln);
        }
        var bands = TextLayoutEngine.BandLines(nonEmpty, opts, consumeWork, cancellationCheck);
        // Fill detailed geometry first
        foreach (var ln in lines) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            page.LinesDetailed.Add(ToStructuredLine(ln));
        }
        // Then semantic classification
        foreach (var ln in lines) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            string t = ln.Text.Trim();
            if (t.Length == 0) continue;
            page.Lines.Add(t);
            if (TryParseTocRow(t, out string tocLabel, out int num)) {
                var label = NormalizeShattered(tocLabel.TrimEnd('.').Trim());
                page.Toc.Add((label, num));
                AddLeaderRow(page, label, num.ToString(System.Globalization.CultureInfo.InvariantCulture));
                continue;
            }
            if (TryParseListItemText(t, out string marker, out string listText, out int listLevel)) {
                page.ListItems.Add(t);
                page.ListNodes.Add(new StructuredListItem {
                    Level = listLevel,
                    Marker = marker,
                    Text = listText,
                    Line = ToStructuredLine(ln)
                });
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
            cancellationCheck?.Invoke();
            if (b.Count == 0) continue;
            consumeWork?.Invoke(b.Count);
            double top = b[0].Y; double bottom = b[b.Count - 1].Y;
            var sb = new StructuredBand { YTop = top, YBottom = bottom };
            foreach (var ln in b) sb.Lines.Add(ln.Text);
            page.Bands.Add(sb);
        }

        // Table detection: prefer banded column inference; fallback to per-line
        var tables = precomputedTables is null
            ? TableDetector.DetectTablesFromBands(
                bands,
                pageHeight,
                consumeWork,
                cancellationCheck)
            : precomputedTables.ToList();
        if (tables.Count > 0) {
            // Clean leaders and add
            foreach (var t in tables) {
                NormalizeDetectedTable(t);
                if (string.Equals(t.Kind, "leaders", StringComparison.OrdinalIgnoreCase)) {
                    // add only to detailed + LeaderRows; do NOT mix into generic Tables
                    page.TablesDetailed.Add(t);
                    foreach (var r in t.Rows) AddLeaderRow(page, r[0], r[1]);
                    continue;
                }
                page.TablesDetailed.Add(t);
                page.Tables.AddRange(t.Rows);
            }
        } else if (precomputedTables is null) {
            // Try a page-level leader-based table (TOC-like)
            var leaderTbl = TableDetector.DetectLeaderTable(
                nonEmpty,
                pageHeight,
                consumeWork,
                cancellationCheck);
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
                var rows = TableDetector.DetectLineRows(
                    lines,
                    pageHeight,
                    consumeWork,
                    cancellationCheck);
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
        cancellationCheck?.Invoke();
        IReadOnlyList<StructuredTableBounds> tableBounds = BuildTableBounds(
            page.TablesDetailed,
            consumeWork,
            cancellationCheck);
        AddHeadings(page, nonEmpty, tableBounds, consumeWork, cancellationCheck);
        cancellationCheck?.Invoke();
        AddParagraphs(page, nonEmpty, fallbackTableLines, tableBounds, consumeWork, cancellationCheck);
        cancellationCheck?.Invoke();
        return page;
    }

    private static void AddParagraphs(
        StructuredPage page,
        List<TextLayoutEngine.TextLine> lines,
        HashSet<TextLayoutEngine.TextLine> fallbackTableLines,
        IReadOnlyList<StructuredTableBounds> tableBounds,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var candidates = new List<TextLayoutEngine.TextLine>();
        foreach (var line in lines) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            string text = line.Text.Trim();
            if (text.Length == 0 ||
                IsListItemText(text) ||
                IsHeadingLine(line, page.Headings, consumeWork, cancellationCheck) ||
                fallbackTableLines.Contains(line) ||
                IsInsideTable(line, tableBounds, consumeWork, cancellationCheck)) {
                continue;
            }

            candidates.Add(line);
        }

        if (candidates.Count == 0) {
            return;
        }

        var gaps = new List<double>();
        for (int i = 1; i < candidates.Count; i++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
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
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
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
                Spans = Array.AsReadOnly(line.Spans.ToArray())
            });
        }

        return paragraph;
    }

    private static void AddHeadings(
        StructuredPage page,
        List<TextLayoutEngine.TextLine> lines,
        IReadOnlyList<StructuredTableBounds> tableBounds,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var bodySizes = new List<double>();
        foreach (var line in lines) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            string text = line.Text.Trim();
            if (text.Length == 0 ||
                IsListItemText(text) ||
                IsInsideTable(line, tableBounds, consumeWork, cancellationCheck)) {
                continue;
            }

            bodySizes.Add(GetLineFontSize(line));
        }

        double bodySize = EstimateBodyFontSize(bodySizes);
        if (bodySize <= 0) {
            return;
        }

        foreach (var line in lines) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            string text = line.Text.Trim();
            if (text.Length == 0 ||
                PdfUnicodeScalarAnalysis.CountScalars(text) > 160 ||
                IsListItemText(text) ||
                IsInsideTable(line, tableBounds, consumeWork, cancellationCheck)) {
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
                Spans = Array.AsReadOnly(line.Spans.ToArray())
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

    private static bool IsHeadingLine(
        TextLayoutEngine.TextLine line,
        List<StructuredHeading> headings,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        for (int i = 0; i < headings.Count; i++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
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
            Spans = Array.AsReadOnly(line.Spans.ToArray())
        };
    }

    internal static bool IsInsideTable(TextLayoutEngine.TextLine line, IReadOnlyList<StructuredTable> tables) =>
        IsInsideTable(line, tables, consumeWork: null, cancellationCheck: null);

    internal static bool IsInsideTable(
        TextLayoutEngine.TextLine line,
        IReadOnlyList<StructuredTable> tables,
        Action<long>? consumeWork,
        Action? cancellationCheck) =>
        IsInsideTable(line, BuildTableBounds(tables, consumeWork, cancellationCheck), consumeWork, cancellationCheck);

    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredTableBounds> BuildTableBounds(
        IReadOnlyList<StructuredTable> tables,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var result = new List<StructuredTableBounds>(tables.Count);
        for (int i = 0; i < tables.Count; i++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            var table = tables[i];
            if (table.Columns.Count == 0) {
                continue;
            }

            double tableLeft = double.MaxValue;
            double tableRight = double.MinValue;
            for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1);
                StructuredTableColumn column = table.Columns[columnIndex];
                tableLeft = Math.Min(tableLeft, Math.Min(column.From, column.To));
                tableRight = Math.Max(tableRight, Math.Max(column.From, column.To));
            }
            result.Add(new StructuredTableBounds(
                Math.Max(table.YTop, table.YBottom),
                Math.Min(table.YTop, table.YBottom),
                tableLeft,
                tableRight));
        }

        return result.AsReadOnly();
    }

    private static bool IsInsideTable(
        TextLayoutEngine.TextLine line,
        IReadOnlyList<StructuredTableBounds> tables,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        double lineLeft = Math.Min(line.XStart, line.XEnd);
        double lineRight = Math.Max(line.XStart, line.XEnd);
        for (int i = 0; i < tables.Count; i++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            StructuredTableBounds table = tables[i];
            if (line.Y <= table.Top + 0.001D &&
                line.Y >= table.Bottom - 0.001D &&
                lineRight >= table.Left - 2D &&
                lineLeft <= table.Right + 2D) return true;
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

    internal static void NormalizeDetectedTable(StructuredTable table) {
        if (string.Equals(table.Kind, "leaders", StringComparison.OrdinalIgnoreCase)) {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                string[] row = table.Rows[rowIndex];
                if (row.Length < 2) continue;
                row[0] = NormalizeShattered(row[0]);
                row[1] = NormalizeLeaderValue(row[1]);
            }
            return;
        }

        // Preserve punctuation and decoded word boundaries. Structural detection must not rewrite
        // cell content such as version identifiers, decimal values, ellipses, or abbreviations.
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            string[] row = table.Rows[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Length; columnIndex++) {
                row[columnIndex] = NormalizeShattered(row[columnIndex]);
            }
        }
    }

    internal static void AddLeaderRow(StructuredPage page, string label, string value) {
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
            int digitCount = 0;
            while (digitEnd < trailingContentEnd && digitCount < 6 &&
                   TryGetDecimalDigit(text, digitEnd, out _, out int consumed)) {
                digitEnd += consumed;
                digitCount++;
            }

            if (digitCount is < 1 or > 5 || digitEnd != trailingContentEnd) {
                continue;
            }

            label = text.Substring(0, runStart);
            for (int digitIndex = digitStart; digitIndex < digitEnd;) {
                if (!TryGetDecimalDigit(text, digitIndex, out int digit, out int consumed)) return false;
                pageNumber = checked((pageNumber * 10) + digit);
                digitIndex += consumed;
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
                !IsUnicodeLetterOrDigit(text, valueStart) ||
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
        int suffixStart = 0;
        for (int index = 0; index < text.Length;) {
            int scalarLength = char.IsSurrogatePair(text, index) ? 2 : 1;
            char character = text[index];
            bool allowed = IsUnicodeLetterOrDigit(text, index) ||
                           char.IsWhiteSpace(character) ||
                           character is '.' or ',' or '\'' or '/' or '%' or '+' or '-' or '(' or ')';
            if (!allowed) suffixStart = index + scalarLength;
            index += scalarLength;
        }
        return suffixStart;
    }

    private static int SkipWhitespace(string text, int index) {
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }

        return index;
    }

    private static bool TryGetDecimalDigit(
        string text,
        int index,
        out int digit,
        out int consumed) {
        digit = System.Globalization.CharUnicodeInfo.GetDecimalDigitValue(text, index);
        consumed = index + 1 < text.Length && char.IsSurrogatePair(text, index) ? 2 : 1;
        return digit >= 0;
    }

    private static bool IsUnicodeLetterOrDigit(string text, int index) {
        if (System.Globalization.CharUnicodeInfo.GetDecimalDigitValue(text, index) >= 0) return true;
        System.Globalization.UnicodeCategory category =
            System.Globalization.CharUnicodeInfo.GetUnicodeCategory(text, index);
        return category is
            System.Globalization.UnicodeCategory.UppercaseLetter or
            System.Globalization.UnicodeCategory.LowercaseLetter or
            System.Globalization.UnicodeCategory.TitlecaseLetter or
            System.Globalization.UnicodeCategory.ModifierLetter or
            System.Globalization.UnicodeCategory.OtherLetter or
            System.Globalization.UnicodeCategory.LetterNumber or
            System.Globalization.UnicodeCategory.OtherNumber;
    }

    private static bool IsLeaderCurrency(char value) =>
        char.GetUnicodeCategory(value) == System.Globalization.UnicodeCategory.CurrencySymbol;

    internal static string NormalizeShattered(string s) {
        if (string.IsNullOrEmpty(s)) return s;
        // Retain decoded word boundaries. Only the positioned-content stage can decide whether
        // separately painted runs are adjacent glyphs or distinct words.
        return System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
    }

    private static string NormalizeLeaderValue(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        string normalized = Regex.Replace(value!.Trim(), "\\s+", " ");

        bool hasDigit = false;
        for (int index = 0; index < normalized.Length;) {
            if (TryGetDecimalDigit(normalized, index, out _, out int consumed)) {
                hasDigit = true;
                break;
            }
            index += char.IsSurrogatePair(normalized, index) ? 2 : 1;
        }

        return hasDigit ? normalized : string.Empty;
    }
}
