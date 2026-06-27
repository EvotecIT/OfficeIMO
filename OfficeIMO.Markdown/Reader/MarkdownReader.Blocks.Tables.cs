using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Determines whether a line is likely to be part of a markdown table. The logic follows
    /// CommonMark's relaxed table rules: when both outer pipes are present a single column is
    /// permitted, otherwise at least two pipe-separated cells are required so that plain
    /// paragraphs containing a single <c>|</c> are not mis-classified as tables.
    /// </summary>
    private static bool LooksLikeTableRow(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.Trim();
        if (trimmed.Length < 3 || !trimmed.Contains('|')) return false;

        var cells = SplitTableRow(trimmed);
        if (cells.Count == 0) return false;

        bool hasLeadingPipe = trimmed[0] == '|';
        bool hasTrailingPipe = trimmed[trimmed.Length - 1] == '|';
        bool hasOuterPipes = hasLeadingPipe && hasTrailingPipe;

        if (!hasOuterPipes && cells.Count < 2) return false;

        return true;
    }

    private static bool StartsTable(string[] lines, int index, MarkdownReaderOptions options) =>
        TryGetTableExtent(lines, index, out _, out _, allowHeaderlessTables: options?.AllowHeaderlessTables ?? true);

    private static bool TryGetTableExtent(
        string[] lines,
        int start,
        out int end,
        out bool hasOuterPipes,
        bool allowSingleRowHeaderless = false,
        bool allowHeaderlessTables = true) {
        end = start;
        hasOuterPipes = false;
        if (lines is null || start < 0 || start >= lines.Length) return false;
        if (!LooksLikeTableRow(lines[start])) return false;
        if (IsAlignmentRow(lines[start])) return false;

        var firstTrimmed = lines[start].Trim();
        if (firstTrimmed.Length == 0) return false;

        bool hasLeadingPipe = firstTrimmed[0] == '|';
        bool hasTrailingPipe = firstTrimmed[firstTrimmed.Length - 1] == '|';
        hasOuterPipes = hasLeadingPipe && hasTrailingPipe;

        int j = start + 1;
        bool sawAlignmentRow = false;
        if (j < lines.Length && IsAlignmentRow(lines[j])) {
            sawAlignmentRow = true;
            var headerCells = SplitTableRow(lines[start]);
            var alignmentCells = SplitTableRow(lines[j]);
            if (headerCells.Count != alignmentCells.Count) {
                return false;
            }

            j++;
        }

        if (!sawAlignmentRow) {
            if (!allowHeaderlessTables && !allowSingleRowHeaderless) {
                return false;
            }

            // Headerless tables are easy to mis-detect (any two lines with pipes). To reduce false positives,
            // require explicit outer pipes on every row and at least two rows.
            if (!hasOuterPipes) return false;

            if (j >= lines.Length) {
                if (allowSingleRowHeaderless) {
                    end = start;
                    return true;
                }

                return false;
            }

            // Require the 2nd row to also have outer pipes, otherwise treat the first row as a paragraph line.
            var second = (lines[j] ?? string.Empty).Trim();
            if (second.Length == 0) {
                if (allowSingleRowHeaderless) {
                    end = start;
                    return true;
                }

                return false;
            }

            if (!(second.Length > 0 && second[0] == '|' && second[second.Length - 1] == '|')) return false;

            while (j < lines.Length) {
                var t = (lines[j] ?? string.Empty).Trim();
                if (t.Length == 0) break;
                if (!LooksLikeTableRow(t)) break;
                if (!(t[0] == '|' && t[t.Length - 1] == '|')) break;
                j++;
            }
        } else {
            while (j < lines.Length && LooksLikeTableBodyRow(lines[j])) j++;
        }

        end = j - 1;
        return true;
    }

    private static TableBlock ParseTable(string[] lines, int start, int end, MarkdownReaderOptions options, MarkdownReaderState state) {
        int headerLine = state.SourceLineOffset + start + 1;
        var cells0 = SplitTableRowWithSourceInfo(lines[start], headerLine, state);
        var table = new TableBlock();
        var inlineOptions = CloneOptionsWithoutFrontMatter(options);
        var inlineState = CloneState(state);
        table.InlineRenderOptions = inlineOptions;
        table.InlineRenderState = inlineState;
        if (start + 1 <= end && IsAlignmentRow(lines[start + 1])) {
            table.Headers.AddRange(cells0.Select(cell => cell.Text));
            table.UseHeaderColumnCountForRendering = true;
            table.SetAlignmentRowSourceSpan(CreateLineSpan(state, state.SourceLineOffset + start + 2, state.SourceLineOffset + start + 2));
            var aligns = SplitTableRow(lines[start + 1]);
            for (int i = 0; i < aligns.Count; i++) table.Alignments.Add(ParseAlignmentCell(aligns[i]));
            for (int i = start + 2; i <= end; i++) {
                int absoluteLine = state.SourceLineOffset + i + 1;
                var row = SplitTableRowWithSourceInfo(lines[i], absoluteLine, state);
                table.Rows.Add(row.Select(cell => cell.Text).ToArray());
            }
        } else {
            for (int i = start; i <= end; i++) {
                int absoluteLine = state.SourceLineOffset + i + 1;
                var row = SplitTableRowWithSourceInfo(lines[i], absoluteLine, state);
                table.Rows.Add(row.Select(cell => cell.Text).ToArray());
            }
        }

        var headerCells = table.Headers.Count > 0
            ? SplitTableRowWithSourceInfo(lines[start], headerLine, state)
            : new List<TableCellSourceFragment>();
        var bodyRows = new List<IReadOnlyList<TableCellSourceFragment>>(table.Rows.Count);
        if (table.Headers.Count > 0) {
            for (int i = start + 2; i <= end; i++) {
                int absoluteLine = state.SourceLineOffset + i + 1;
                bodyRows.Add(SplitTableRowWithSourceInfo(lines[i], absoluteLine, state));
            }
        } else {
            for (int i = start; i <= end; i++) {
                int absoluteLine = state.SourceLineOffset + i + 1;
                bodyRows.Add(SplitTableRowWithSourceInfo(lines[i], absoluteLine, state));
            }
        }

        table.SetParsedCells(
            ParseTableInlineCells(headerCells, inlineOptions, inlineState),
            ParseTableInlineRows(bodyRows, inlineOptions, inlineState),
            table.ComputeContentSignature());
        table.SetStructuredCells(
            BuildTableCells(headerCells, inlineOptions, inlineState),
            BuildTableRows(bodyRows, inlineOptions, inlineState),
            table.ComputeContentSignature());
        return table;
    }

    private static List<InlineSequence> ParseTableInlineCells(IReadOnlyList<TableCellSourceFragment> cells, MarkdownReaderOptions options, MarkdownReaderState state) {
        var parsedCells = new List<InlineSequence>(cells?.Count ?? 0);
        if (cells == null) {
            return parsedCells;
        }

        for (int i = 0; i < cells.Count; i++) {
            parsedCells.Add(ParseTableCellInlines(cells[i], options, state));
        }
        return parsedCells;
    }

    private static List<IReadOnlyList<InlineSequence>> ParseTableInlineRows(IReadOnlyList<IReadOnlyList<TableCellSourceFragment>> rows, MarkdownReaderOptions options, MarkdownReaderState state) {
        var parsedRows = new List<IReadOnlyList<InlineSequence>>(rows.Count);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            if (row == null || row.Count == 0) {
                parsedRows.Add(Array.Empty<InlineSequence>());
                continue;
            }

            var parsedRow = new List<InlineSequence>(row.Count);
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                parsedRow.Add(ParseTableCellInlines(row[cellIndex], options, state));
            }
            parsedRows.Add(parsedRow);
        }
        return parsedRows;
    }

    private static InlineSequence ParseTableCellInlines(TableCellSourceFragment cell, MarkdownReaderOptions options, MarkdownReaderState state) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return new InlineSequence();
        }

        var (sanitized, sourceMap) = NormalizeTableCellInlineMarkdownWithSourceMap(cell, state);
        return ParseInlines(sanitized, options, state, sourceMap);
    }

    private static List<TableCell> BuildTableCells(IReadOnlyList<TableCellSourceFragment> cells, MarkdownReaderOptions options, MarkdownReaderState state) {
        var typedCells = new List<TableCell>(cells?.Count ?? 0);
        if (cells == null) {
            return typedCells;
        }

        for (int i = 0; i < cells.Count; i++) {
            typedCells.Add(BuildTableCell(cells[i], options, state));
        }
        return typedCells;
    }

    private static List<IReadOnlyList<TableCell>> BuildTableRows(IReadOnlyList<IReadOnlyList<TableCellSourceFragment>> rows, MarkdownReaderOptions options, MarkdownReaderState state) {
        var typedRows = new List<IReadOnlyList<TableCell>>(rows.Count);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            if (row == null || row.Count == 0) {
                typedRows.Add(Array.Empty<TableCell>());
                continue;
            }

            typedRows.Add(BuildTableCells(row, options, state));
        }
        return typedRows;
    }

    private static TableCell BuildTableCell(TableCellSourceFragment cell, MarkdownReaderOptions options, MarkdownReaderState state) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return new TableCell() {
                SourceSpan = cell.SourceSpan
            };
        }

        var structuredCell = TryParseStructuredTableCellBlocks(cell, options, state);
        if (structuredCell.HasValue) {
            return new TableCell(structuredCell.Value.Blocks) {
                SourceSpan = cell.SourceSpan,
                SyntaxChildren = structuredCell.Value.SyntaxChildren
            };
        }

        return new TableCell(new[] {
            new ParagraphBlock(ParseTableCellInlines(cell, options, state))
        }) {
            SourceSpan = cell.SourceSpan
        };
    }

    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren)? TryParseStructuredTableCellBlocks(
        TableCellSourceFragment cell,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return null;
        }
        if (!options.ParseTableCellBlocks) {
            return null;
        }

        var normalized = BuildTableCellSourceLines(cell);
        if (!TableBlock.LooksLikeStructuredMarkdownCell(string.Join("\n", normalized.Select(line => line.Text)))) {
            return null;
        }

        var fragmentOptions = CloneOptionsWithoutFrontMatter(options);
        fragmentOptions.Tables = false;
        var fragmentState = CloneState(state);
        var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(normalized, fragmentOptions, fragmentState);
        if (blocks.Count == 0) {
            return null;
        }

        if (TableBlock.ContainsUnsafeRawHtmlTableCellBlocks(blocks)) {
            return null;
        }

        if (blocks.Count == 1 && blocks[0] is ParagraphBlock) {
            return null;
        }

        return (blocks, syntaxChildren);
    }

    private readonly struct TableCellSourceFragment {
        public TableCellSourceFragment(string markdown, string text, MarkdownSourceSpan sourceSpan) {
            Markdown = markdown ?? string.Empty;
            Text = text ?? string.Empty;
            SourceSpan = sourceSpan;
        }

        public string Markdown { get; }
        public string Text { get; }
        public MarkdownSourceSpan SourceSpan { get; }
    }

    private static bool IsAlignmentRow(string line) {
        var cells = SplitTableRow(line);
        if (cells.Count == 0) return false;
        foreach (var c in cells) {
            var t = c.Trim(); if (t.Length == 0) return false;
            int dash = 0;
            for (int i = 0; i < t.Length; i++) {
                char ch = t[i];
                if (ch == '-') dash++;
                else if (ch == ':' && (i == 0 || i == t.Length - 1)) { } else return false;
            }
            if (dash < 1) return false;
        }
        return true;
    }

    private static bool LooksLikeTableBodyRow(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        return line.Trim().Contains('|');
    }

    private static ColumnAlignment ParseAlignmentCell(string cell) {
        var t = cell.Trim();
        if (t.StartsWith(":")) { if (t.EndsWith(":")) return ColumnAlignment.Center; return ColumnAlignment.Left; }
        if (t.EndsWith(":")) return ColumnAlignment.Right;
        return ColumnAlignment.None;
    }

    private static List<string> SplitTableRow(string line) {
        if (line is null) return new List<string>();
        return SplitTableRowWithSourceInfo(line, absoluteLine: 1, state: null).Select(cell => cell.Text).ToList();
    }

    private static List<TableCellSourceFragment> SplitTableRowWithSourceInfo(string line, int absoluteLine, MarkdownReaderState? state) {
        if (line is null) {
            return new List<TableCellSourceFragment>();
        }

        int trimStart = 0;
        int trimEndExclusive = line.Length;
        while (trimStart < trimEndExclusive && char.IsWhiteSpace(line[trimStart])) {
            trimStart++;
        }
        while (trimEndExclusive > trimStart && char.IsWhiteSpace(line[trimEndExclusive - 1])) {
            trimEndExclusive--;
        }

        int contentStart = trimStart;
        int contentEndExclusive = trimEndExclusive;
        if (contentStart < contentEndExclusive && line[contentStart] == '|') {
            contentStart++;
        }
        if (contentEndExclusive > contentStart && line[contentEndExclusive - 1] == '|') {
            contentEndExclusive--;
        }

        var cells = new List<TableCellSourceFragment>();
        int segmentStart = contentStart;
        int index = contentStart;
        int codeFenceLen = 0;

        while (index < contentEndExclusive) {
            char ch = line[index];
            if (ch == '\\' && index + 1 < contentEndExclusive) {
                index += 2;
                continue;
            }

            if (ch == '`') {
                int run = 1;
                int tick = index + 1;
                while (tick < contentEndExclusive && line[tick] == '`') {
                    run++;
                    tick++;
                }

                if (codeFenceLen == 0) {
                    if (HasClosingBacktickRun(line.Substring(contentStart, contentEndExclusive - contentStart), index - contentStart + run, run)) {
                        codeFenceLen = run;
                    }
                } else if (run == codeFenceLen) {
                    codeFenceLen = 0;
                }

                index += run;
                continue;
            }

            if (ch == '|' && codeFenceLen == 0) {
                cells.Add(CreateTableCellSourceFragment(line, segmentStart, index, absoluteLine, state));
                segmentStart = index + 1;
            }

            index++;
        }

        cells.Add(CreateTableCellSourceFragment(line, segmentStart, contentEndExclusive, absoluteLine, state));
        return cells;
    }

    private static TableCellSourceFragment CreateTableCellSourceFragment(
        string line,
        int segmentStart,
        int segmentEndExclusive,
        int absoluteLine,
        MarkdownReaderState? state) {
        int trimmedStart = segmentStart;
        int trimmedEndExclusive = segmentEndExclusive;

        while (trimmedStart < trimmedEndExclusive && char.IsWhiteSpace(line[trimmedStart])) {
            trimmedStart++;
        }

        while (trimmedEndExclusive > trimmedStart && char.IsWhiteSpace(line[trimmedEndExclusive - 1])) {
            trimmedEndExclusive--;
        }

        bool hasContent = trimmedStart < trimmedEndExclusive;
        string markdown = hasContent
            ? line.Substring(trimmedStart, trimmedEndExclusive - trimmedStart)
            : string.Empty;
        string text = UnescapeBackslashEscapesOutsideCodeSpans(markdown);
        int startColumn = hasContent ? trimmedStart + 1 : segmentStart + 1;
        int endColumn = hasContent ? trimmedEndExclusive : Math.Max(startColumn, segmentEndExclusive);
        var span = CreateSpan(state, absoluteLine, startColumn, absoluteLine, endColumn);
        return new TableCellSourceFragment(markdown, text, span);
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) NormalizeTableCellInlineMarkdownWithSourceMap(
        TableCellSourceFragment cell,
        MarkdownReaderState? state) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return (string.Empty, null);
        }

        var builder = new StringBuilder(cell.Markdown.Length + 8);
        List<MarkdownSourcePoint?>? points = state?.SourceTextMap == null ? null : new List<MarkdownSourcePoint?>(cell.Markdown.Length + 8);
        int absoluteLine = cell.SourceSpan.StartLine;
        int column = cell.SourceSpan.StartColumn ?? 1;
        int codeFenceLen = 0;

        void AppendMapped(string value, int sourceColumn) {
            builder.Append(value);
            if (points == null) {
                return;
            }

            var point = state!.SourceTextMap!.CreatePoint(absoluteLine, sourceColumn);
            for (int i = 0; i < value.Length; i++) {
                points.Add(point);
            }
        }

        for (int i = 0; i < cell.Markdown.Length; i++) {
            char ch = cell.Markdown[i];
            if (ch == '`') {
                int run = 1;
                int tick = i + 1;
                while (tick < cell.Markdown.Length && cell.Markdown[tick] == '`') {
                    run++;
                    tick++;
                }

                if (codeFenceLen == 0) {
                    if (HasClosingBacktickRun(cell.Markdown, i + run, run)) {
                        codeFenceLen = run;
                    }
                } else if (run == codeFenceLen) {
                    codeFenceLen = 0;
                }

                AppendMapped(cell.Markdown.Substring(i, run), column);
                column += run;
                i += run - 1;
                continue;
            }

            if (codeFenceLen != 0 && ch == '\\' && i + 1 < cell.Markdown.Length && cell.Markdown[i + 1] == '|') {
                AppendMapped("|", column + 1);
                column += 2;
                i++;
                continue;
            }

            if (ch == '\r') {
                if (i + 1 < cell.Markdown.Length && cell.Markdown[i + 1] == '\n') {
                    i++;
                }
                AppendMapped("\n", column);
                column++;
                continue;
            }

            if (ch == '\n') {
                AppendMapped("\n", column);
                column++;
                continue;
            }

            if (ch == '<' && TableBlock.TryConsumeBreakTag(cell.Markdown, i, out int consumed)) {
                AppendMapped("\n", column);
                column += consumed;
                i += consumed - 1;
                continue;
            }

            switch (ch) {
                case '<':
                    if (TableBlock.TryConsumeSupportedInlineFormattingTag(cell.Markdown, i, out int formattingTagLength)) {
                        AppendMapped(cell.Markdown.Substring(i, formattingTagLength), column);
                        column += formattingTagLength;
                        i += formattingTagLength - 1;
                        continue;
                    }

                    if (TryConsumeInlineRawHtmlTableCellTag(cell.Markdown, i, out int rawHtmlTagLength)) {
                        AppendMapped(cell.Markdown.Substring(i, rawHtmlTagLength), column);
                        column += rawHtmlTagLength;
                        i += rawHtmlTagLength - 1;
                        continue;
                    }

                    AppendMapped("&lt;", column);
                    break;
                case '>':
                    AppendMapped("&gt;", column);
                    break;
                case '&':
                    AppendMapped("&amp;", column);
                    break;
                default:
                    AppendMapped(ch.ToString(), column);
                    break;
            }

            column++;
        }

        return (builder.ToString(), points == null ? null : new MarkdownInlineSourceMap(points.ToArray()));
    }

    private static bool TryConsumeInlineRawHtmlTableCellTag(string markdown, int index, out int consumed) {
        consumed = 0;
        if (!TryConsumeRawInlineHtmlTag(markdown, index, out int rawHtmlTagLength)) {
            return false;
        }

        string rawTag = markdown.Substring(index, rawHtmlTagLength);
        if (!HtmlBlockParser.TryParseHtmlTag(rawTag, out var tagName, out _, out _)) {
            return false;
        }

        if (HtmlBlockParser.IsBlockOrRawTextHtmlTagName(tagName)) {
            return false;
        }

        consumed = rawHtmlTagLength;
        return true;
    }

    private static List<MarkdownSourceLineSlice> BuildTableCellSourceLines(TableCellSourceFragment cell) {
        var slices = new List<MarkdownSourceLineSlice>();
        if (string.IsNullOrEmpty(cell.Markdown)) {
            slices.Add(new MarkdownSourceLineSlice(string.Empty, cell.SourceSpan.StartLine, cell.SourceSpan.StartColumn ?? 1));
            return slices;
        }

        var current = new StringBuilder(cell.Markdown.Length);
        int absoluteLine = cell.SourceSpan.StartLine;
        int column = cell.SourceSpan.StartColumn ?? 1;
        int currentStartColumn = column;

        void Flush() {
            slices.Add(new MarkdownSourceLineSlice(current.ToString(), absoluteLine, currentStartColumn));
            current.Clear();
            currentStartColumn = column;
        }

        for (int i = 0; i < cell.Markdown.Length; i++) {
            char ch = cell.Markdown[i];
            if (ch == '\r') {
                if (i + 1 < cell.Markdown.Length && cell.Markdown[i + 1] == '\n') {
                    i++;
                }
                Flush();
                column++;
                currentStartColumn = column;
                continue;
            }

            if (ch == '\n') {
                Flush();
                column++;
                currentStartColumn = column;
                continue;
            }

            if (ch == '<' && TableBlock.TryConsumeBreakTag(cell.Markdown, i, out int consumed)) {
                Flush();
                column += consumed;
                currentStartColumn = column;
                i += consumed - 1;
                continue;
            }

            current.Append(ch);
            column++;
        }

        Flush();
        return slices;
    }

    private static string UnescapeBackslashEscapesOutsideCodeSpans(string value) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;

        var sb = new StringBuilder(value.Length);
        int i = 0;
        int codeFenceLen = 0;

        while (i < value.Length) {
            char ch = value[i];

            if (ch == '`') {
                int run = 1;
                int j = i + 1;
                while (j < value.Length && value[j] == '`') { run++; j++; }

                if (codeFenceLen == 0) {
                    if (HasClosingBacktickRun(value, j, run)) {
                        codeFenceLen = run;
                    }
                }
                else if (run == codeFenceLen) codeFenceLen = 0;

                sb.Append(value, i, run);
                i += run;
                continue;
            }

            if (ch == '\\' && codeFenceLen == 0 && i + 1 < value.Length) {
                char next = value[i + 1];
                if (IsBackslashEscapable(next)) {
                    sb.Append(next);
                    i += 2;
                    continue;
                }
            }

            sb.Append(ch);
            i++;
        }

        return sb.ToString();
    }

    private static bool HasClosingBacktickRun(string text, int start, int runLength) {
        if (string.IsNullOrEmpty(text) || start >= text.Length) return false;

        for (int i = start; i < text.Length; i++) {
            if (text[i] != '`') continue;

            int run = 1;
            int j = i + 1;
            while (j < text.Length && text[j] == '`') {
                run++;
                j++;
            }

            if (run == runLength) return true;
            i = j - 1;
        }

        return false;
    }

}
