using System;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Pipe table with optional header row.
/// </summary>
public sealed class TableBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock, IChildMarkdownBlockContainer {
    private IReadOnlyList<TableCell>? _cachedHeaderCells;
    private IReadOnlyList<IReadOnlyList<TableCell>>? _cachedRowCells;
    private int? _cachedCellContentSignature;
    private bool _cachedUsesStructuredCells;
    private int _cachedCellColumnCount = -1;

    /// <summary>Optional header cells.</summary>
    public List<string> Headers { get; } = new List<string>();
    /// <summary>Typed header cell content.</summary>
    public IReadOnlyList<TableCell> HeaderCells => GetOrBuildHeaderCells();
    /// <summary>Parsed inline representation of the current header cells.</summary>
    public IReadOnlyList<InlineSequence> HeaderInlines => BuildHeaderInlines();
    /// <summary>Data rows.</summary>
    public List<IReadOnlyList<string>> Rows { get; } = new List<IReadOnlyList<string>>();
    /// <summary>Typed row cell content.</summary>
    public IReadOnlyList<IReadOnlyList<TableCell>> RowCells => GetOrBuildRowCells();
    /// <summary>Enumerates header and body cells in document order, preserving row/column metadata on each cell.</summary>
    public IEnumerable<TableCell> EnumerateCells() {
        var headers = HeaderCells;
        for (int i = 0; i < headers.Count; i++) {
            yield return headers[i];
        }

        var rows = RowCells;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                yield return row[columnIndex];
            }
        }
    }

    /// <summary>Gets a body cell by zero-based row and column index.</summary>
    public TableCell? GetCell(int rowIndex, int columnIndex) {
        if (rowIndex < 0 || columnIndex < 0) {
            return null;
        }

        var rows = RowCells;
        if (rowIndex >= rows.Count) {
            return null;
        }

        var row = rows[rowIndex];
        return columnIndex < row.Count ? row[columnIndex] : null;
    }

    /// <summary>Gets a header cell by zero-based column index.</summary>
    public TableCell? GetHeaderCell(int columnIndex) {
        if (columnIndex < 0) {
            return null;
        }

        var headers = HeaderCells;
        return columnIndex < headers.Count ? headers[columnIndex] : null;
    }
    /// <summary>Parsed inline representation of the current data rows.</summary>
    public IReadOnlyList<IReadOnlyList<InlineSequence>> RowInlines => BuildRowInlines();
    /// <summary>Optional column alignments per column (used when headers are present).</summary>
    public List<ColumnAlignment> Alignments { get; } = new List<ColumnAlignment>();
    /// <summary>Number of rows skipped due to table limits.</summary>
    public int SkippedRowCount { get; internal set; }
    /// <summary>Number of columns skipped due to table limits.</summary>
    public int SkippedColumnCount { get; internal set; }
    internal List<InlineSequence>? ParsedHeaders { get; private set; }
    internal List<IReadOnlyList<InlineSequence>>? ParsedRows { get; private set; }
    internal int? ParsedContentSignature { get; private set; }
    internal List<TableCell>? StructuredHeaders { get; private set; }
    internal List<IReadOnlyList<TableCell>>? StructuredRows { get; private set; }
    internal int? StructuredContentSignature { get; private set; }

    // When a table is produced by the reader, we keep the parse options/state so inline parsing in cells
    // (links/emphasis/etc) can honor URL safety settings and reference-style link definitions.
    internal MarkdownReaderOptions? InlineRenderOptions { get; set; }
    internal MarkdownReaderState? InlineRenderState { get; set; }

    internal void SetParsedCells(
        IReadOnlyList<InlineSequence>? headers,
        IReadOnlyList<IReadOnlyList<InlineSequence>>? rows,
        int contentSignature) {
        ParsedHeaders = headers == null ? null : new List<InlineSequence>(headers);
        if (rows == null) {
            ParsedRows = null;
        } else {
            ParsedRows = new List<IReadOnlyList<InlineSequence>>(rows.Count);
            for (int i = 0; i < rows.Count; i++) {
                var row = rows[i];
                ParsedRows.Add(row == null ? Array.Empty<InlineSequence>() : new List<InlineSequence>(row));
            }
        }

        ParsedContentSignature = contentSignature;
        InvalidateRealizedCellCache();
    }

    internal void SetStructuredCells(
        IReadOnlyList<TableCell>? headers,
        IReadOnlyList<IReadOnlyList<TableCell>>? rows,
        int contentSignature) {
        StructuredHeaders = headers == null ? null : CloneStructuredRow(headers);
        if (rows == null) {
            StructuredRows = null;
        } else {
            StructuredRows = new List<IReadOnlyList<TableCell>>(rows.Count);
            for (int i = 0; i < rows.Count; i++) {
                StructuredRows.Add(rows[i] == null ? Array.Empty<TableCell>() : CloneStructuredRow(rows[i]));
            }
        }

        StructuredContentSignature = contentSignature;
        InvalidateRealizedCellCache();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        static void AppendRow(StringBuilder builder, IReadOnlyList<string> cells) {
            builder.Append("| ");
            builder.Append(string.Join(" | ", cells));
            builder.Append(" |\n");
        }

        int columnCount = GetEffectiveColumnCount();
        bool useStructuredCells = StructuredContentSignature.HasValue && StructuredContentSignature.Value == ComputeContentSignature();

        if (Headers.Count > 0) {
            var sb = new StringBuilder();
            var headerMarkdown = useStructuredCells
                ? PrepareStructuredRowMarkdown(StructuredHeaders, Headers, columnCount)
                : PrepareRowCells(Headers, columnCount);
            var escapedHeaders = headerMarkdown.Select(EscapeMarkdownCell).ToArray();
            AppendRow(sb, escapedHeaders);

            var alignRow = new string[columnCount];
            for (int i = 0; i < columnCount; i++) {
                var a = GetAlignment(i);
                alignRow[i] = a switch { ColumnAlignment.Left => ":---", ColumnAlignment.Center => ":---:", ColumnAlignment.Right => "---:", _ => "---" };
            }
            AppendRow(sb, alignRow);

            for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
                var rowMarkdown = useStructuredCells && StructuredRows != null && rowIndex < StructuredRows.Count
                    ? PrepareStructuredRowMarkdown(StructuredRows[rowIndex], Rows[rowIndex], columnCount)
                    : PrepareRowCells(Rows[rowIndex], columnCount);
                var escapedRow = rowMarkdown.Select(EscapeMarkdownCell).ToArray();
                AppendRow(sb, escapedRow);
            }

            return sb.ToString().TrimEnd('\n');
        }

        var sbNoHeaders = new StringBuilder();
        for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
            var rowMarkdown = useStructuredCells && StructuredRows != null && rowIndex < StructuredRows.Count
                ? PrepareStructuredRowMarkdown(StructuredRows[rowIndex], Rows[rowIndex], columnCount)
                : PrepareRowCells(Rows[rowIndex], columnCount);
            var escapedRow = rowMarkdown.Select(EscapeMarkdownCell).ToArray();
            AppendRow(sbNoHeaders, escapedRow);
        }
        return sbNoHeaders.ToString().TrimEnd('\n');
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        StringBuilder sb = new StringBuilder();
        sb.Append("<table>");
        var headerCells = HeaderCells;
        var rowCells = RowCells;
        var headerInlines = BuildHeaderInlines();
        var rowInlines = BuildRowInlines();
        if (Headers.Count > 0) {
            sb.Append("<thead><tr>");
            int columnCount = GetEffectiveColumnCount();
            var preparedHeaders = PrepareRowCells(Headers, columnCount);
            var preparedStructuredHeaders = PrepareStructuredRowCells(headerCells, columnCount);
            var preparedParsedHeaders = PrepareParsedRowCells(headerInlines, columnCount);
            for (int i = 0; i < preparedHeaders.Count; i++) {
                var h = preparedHeaders[i];
                var style = GetAlignment(i);
                var styleAttr = style switch { ColumnAlignment.Left => " style=\"text-align:left\"", ColumnAlignment.Center => " style=\"text-align:center\"", ColumnAlignment.Right => " style=\"text-align:right\"", _ => string.Empty };
                sb.Append($"<th{styleAttr}>");
                sb.Append(RenderCellHtml(h, preparedStructuredHeaders?[i], preparedParsedHeaders?[i]));
                sb.Append("</th>");
            }
            sb.Append("</tr></thead>");
        }
        sb.Append("<tbody>");
        int bodyColumnCount = GetEffectiveColumnCount();
        for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
            var row = Rows[rowIndex];
            var cells = PrepareRowCells(row, bodyColumnCount);
            var structuredCells = rowIndex < rowCells.Count
                ? PrepareStructuredRowCells(rowCells[rowIndex], bodyColumnCount)
                : null;
            var parsedCells = rowIndex < rowInlines.Count
                ? PrepareParsedRowCells(rowInlines[rowIndex], bodyColumnCount)
                : null;
            sb.Append("<tr>");
            for (int i = 0; i < cells.Count; i++) {
                var cell = cells[i];
                var style = GetAlignment(i);
                var styleAttr = style switch { ColumnAlignment.Left => " style=\"text-align:left\"", ColumnAlignment.Center => " style=\"text-align:center\"", ColumnAlignment.Right => " style=\"text-align:right\"", _ => string.Empty };
                sb.Append($"<td{styleAttr}>");
                sb.Append(RenderCellHtml(cell, structuredCells?[i], parsedCells?[i]));
                sb.Append("</td>");
            }
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table>");
        return sb.ToString();
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => BuildChildBlocks();

    private IReadOnlyList<TableCell> GetOrBuildHeaderCells() {
        EnsureRealizedCells();
        return _cachedHeaderCells ?? Array.Empty<TableCell>();
    }

    private IReadOnlyList<IReadOnlyList<TableCell>> GetOrBuildRowCells() {
        EnsureRealizedCells();
        return _cachedRowCells ?? Array.Empty<IReadOnlyList<TableCell>>();
    }

    private void EnsureRealizedCells() {
        int contentSignature = ComputeContentSignature();
        int columnCount = GetEffectiveColumnCount();
        bool useStructuredCells = StructuredContentSignature.HasValue && StructuredContentSignature.Value == contentSignature;

        if (_cachedHeaderCells != null
            && _cachedRowCells != null
            && _cachedCellContentSignature == contentSignature
            && _cachedUsesStructuredCells == useStructuredCells
            && _cachedCellColumnCount == columnCount) {
            return;
        }

        _cachedHeaderCells = BuildHeaderCellsCore(columnCount, useStructuredCells);
        _cachedRowCells = BuildRowCellsCore(columnCount, useStructuredCells);
        _cachedCellContentSignature = contentSignature;
        _cachedUsesStructuredCells = useStructuredCells;
        _cachedCellColumnCount = columnCount;
    }

    private IReadOnlyList<TableCell> BuildHeaderCellsCore(int columnCount, bool useStructuredCells) {
        if (columnCount == 0) {
            return Array.Empty<TableCell>();
        }

        if (useStructuredCells) {
            return AssignTableCellLocations(PrepareStructuredRowCells(StructuredHeaders, columnCount), isHeader: true, rowIndex: -1);
        }

        var headers = PrepareRowCells(Headers, columnCount);
        return AssignTableCellLocations(BuildSimpleRowCells(headers), isHeader: true, rowIndex: -1);
    }

    private IReadOnlyList<IReadOnlyList<TableCell>> BuildRowCellsCore(int columnCount, bool useStructuredCells) {
        if (Rows.Count == 0) {
            return Array.Empty<IReadOnlyList<TableCell>>();
        }

        var rows = new List<IReadOnlyList<TableCell>>(Rows.Count);

        for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
            if (useStructuredCells && StructuredRows != null && rowIndex < StructuredRows.Count) {
                rows.Add(AssignTableCellLocations(PrepareStructuredRowCells(StructuredRows[rowIndex], columnCount), isHeader: false, rowIndex: rowIndex));
                continue;
            }

            rows.Add(AssignTableCellLocations(BuildSimpleRowCells(PrepareRowCells(Rows[rowIndex], columnCount)), isHeader: false, rowIndex: rowIndex));
        }

        return rows;
    }

    private IReadOnlyList<InlineSequence> BuildHeaderInlines() {
        int columnCount = GetEffectiveColumnCount();
        if (columnCount == 0) {
            return Array.Empty<InlineSequence>();
        }

        bool useParsedCells = ParsedContentSignature.HasValue && ParsedContentSignature.Value == ComputeContentSignature();
        if (useParsedCells) {
            return PrepareParsedRowCells(ParsedHeaders, columnCount);
        }

        var headers = PrepareRowCells(Headers, columnCount);
        return ParseInlineCells(headers);
    }

    private IReadOnlyList<IReadOnlyList<InlineSequence>> BuildRowInlines() {
        if (Rows.Count == 0) {
            return Array.Empty<IReadOnlyList<InlineSequence>>();
        }

        int columnCount = GetEffectiveColumnCount();
        bool useParsedCells = ParsedContentSignature.HasValue && ParsedContentSignature.Value == ComputeContentSignature();
        var rows = new List<IReadOnlyList<InlineSequence>>(Rows.Count);

        for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
            if (useParsedCells && ParsedRows != null && rowIndex < ParsedRows.Count) {
                rows.Add(PrepareParsedRowCells(ParsedRows[rowIndex], columnCount));
                continue;
            }

            rows.Add(ParseInlineCells(PrepareRowCells(Rows[rowIndex], columnCount)));
        }

        return rows;
    }

    private IReadOnlyList<InlineSequence> ParseInlineCells(IReadOnlyList<string> cells) {
        if (cells == null || cells.Count == 0) {
            return Array.Empty<InlineSequence>();
        }

        var options = InlineRenderOptions ?? new MarkdownReaderOptions();
        var state = InlineRenderState ?? new MarkdownReaderState();
        var parsed = new InlineSequence[cells.Count];
        for (int i = 0; i < cells.Count; i++) {
            var cell = cells[i];
            if (string.IsNullOrEmpty(cell)) {
                parsed[i] = new InlineSequence();
                continue;
            }

            var normalized = NormalizeBreakMarkers(cell);
            var sanitized = SanitizeInlineMarkdownInput(normalized);
            parsed[i] = MarkdownReader.ParseInlineText(sanitized, options, state);
        }

        return parsed;
    }

    private string RenderCellHtml(string cell, TableCell? structuredCell, InlineSequence? parsedCell) {
        if (structuredCell != null) {
            return NormalizeEncodedEntities(structuredCell.RenderHtml());
        }

        if (parsedCell != null) {
            var parsedRendered = NormalizeEncodedEntities(parsedCell.RenderHtml());
            return parsedRendered.Contains('\n') ? parsedRendered.Replace("\n", "<br/>") : parsedRendered;
        }

        if (string.IsNullOrEmpty(cell)) return string.Empty;
        // Allow simple <br> markers inside table cells and support inline markdown (code, links, emphasis).
        // We avoid allowing arbitrary HTML by translating only <br> tags to hard breaks and then using the inline parser.
        var normalized = NormalizeBreakMarkers(cell);
        var sanitized = SanitizeInlineMarkdownInput(normalized);
        var inlines = MarkdownReader.ParseInlineText(sanitized, InlineRenderOptions, InlineRenderState);
        var rendered = inlines.RenderHtml();
        rendered = NormalizeEncodedEntities(rendered);
        return rendered.Contains('\n') ? rendered.Replace("\n", "<br/>") : rendered;
    }

    private static string EscapeMarkdownCell(string? cell) {
        if (string.IsNullOrEmpty(cell)) return string.Empty;

        var value = cell!;
        StringBuilder? builder = null;

        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '\r':
                    builder ??= AllocateCellBuilder(value, i);
                    if (i + 1 < value.Length && value[i + 1] == '\n') {
                        i++;
                    }
                    builder.Append("<br>");
                    break;
                case '\n':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("<br>");
                    break;
                case '\\':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("\\\\");
                    break;
                case '|':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("\\|");
                    break;
                default:
                    builder?.Append(ch);
                    break;
            }
        }

        return builder?.ToString() ?? value;
    }

    internal static string NormalizeBreakMarkers(string cell) {
        var builder = new StringBuilder(cell.Length);

        for (int i = 0; i < cell.Length; i++) {
            char ch = cell[i];
            switch (ch) {
                case '\r':
                    if (i + 1 < cell.Length && cell[i + 1] == '\n') {
                        i++;
                    }
                    builder.Append('\n');
                    break;
                case '\n':
                    builder.Append('\n');
                    break;
                case '<':
                    if (TryConsumeBreakTag(cell, i, out int consumed)) {
                        builder.Append('\n');
                        i += consumed - 1;
                    } else {
                        builder.Append(ch);
                    }
                    break;
                default:
                    builder.Append(ch);
                    break;
            }
        }

        return builder.ToString();
    }

    internal static string SanitizeInlineMarkdownInput(string value) {
        StringBuilder? builder = null;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '<':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("&lt;");
                    break;
                case '>':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("&gt;");
                    break;
                case '&':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("&amp;");
                    break;
                default:
                    builder?.Append(ch);
                    break;
            }
        }

        return builder?.ToString() ?? value;
    }

    internal static string NormalizeEncodedEntities(string value) {
        if (value.IndexOf("&amp;", StringComparison.Ordinal) < 0) {
            return value;
        }

        return value
            .Replace("&amp;lt;", "&lt;")
            .Replace("&amp;gt;", "&gt;")
            .Replace("&amp;amp;", "&amp;");
    }

    private static StringBuilder AllocateCellBuilder(string seed, int copyLength) {
        var builder = new StringBuilder(seed.Length + 8);
        if (copyLength > 0) {
            builder.Append(seed, 0, copyLength);
        }
        return builder;
    }

    private int GetEffectiveColumnCount() {
        int columnCount = Headers.Count;

        foreach (var row in Rows) {
            if (row != null) {
                columnCount = Math.Max(columnCount, row.Count);
            }
        }

        columnCount = Math.Max(columnCount, Alignments.Count);
        return columnCount;
    }

    private ColumnAlignment GetAlignment(int index) {
        if (index < 0) return ColumnAlignment.None;
        return index < Alignments.Count ? Alignments[index] : ColumnAlignment.None;
    }

    private static IReadOnlyList<string> PrepareRowCells(IReadOnlyList<string>? row, int expectedCount) {
        if (row == null || row.Count == 0) {
            if (expectedCount <= 0) {
                return Array.Empty<string>();
            }

            var padded = new string[expectedCount];
            for (int i = 0; i < padded.Length; i++) {
                padded[i] = string.Empty;
            }
            return padded;
        }

        if (expectedCount <= 0) {
            var copy = new string[row.Count];
            for (int i = 0; i < row.Count; i++) {
                copy[i] = row[i] ?? string.Empty;
            }
            return copy;
        }

        var cells = new string[expectedCount];
        int limit = Math.Min(expectedCount, row.Count);
        for (int i = 0; i < limit; i++) {
            cells[i] = row[i] ?? string.Empty;
        }
        if (limit < expectedCount) {
            for (int i = limit; i < expectedCount; i++) {
                cells[i] = string.Empty;
            }
        }
        return cells;
    }

    private static IReadOnlyList<TableCell> PrepareStructuredRowCells(IReadOnlyList<TableCell>? row, int expectedCount) {
        if (row == null || row.Count == 0) {
            if (expectedCount <= 0) {
                return Array.Empty<TableCell>();
            }

            var padded = new TableCell[expectedCount];
            for (int i = 0; i < padded.Length; i++) {
                padded[i] = new TableCell();
            }
            return padded;
        }

        if (expectedCount <= 0) {
            return CloneStructuredRow(row);
        }

        var cells = new TableCell[expectedCount];
        int limit = Math.Min(expectedCount, row.Count);
        for (int i = 0; i < limit; i++) {
            cells[i] = CloneStructuredCell(row[i]);
        }
        if (limit < expectedCount) {
            for (int i = limit; i < expectedCount; i++) {
                cells[i] = new TableCell();
            }
        }
        return cells;
    }

    private static IReadOnlyList<InlineSequence> PrepareParsedRowCells(IReadOnlyList<InlineSequence>? row, int expectedCount) {
        if (row == null || row.Count == 0) {
            if (expectedCount <= 0) {
                return Array.Empty<InlineSequence>();
            }

            var padded = new InlineSequence[expectedCount];
            for (int i = 0; i < padded.Length; i++) {
                padded[i] = new InlineSequence();
            }
            return padded;
        }

        if (expectedCount <= 0) {
            var copy = new InlineSequence[row.Count];
            for (int i = 0; i < row.Count; i++) {
                copy[i] = row[i] ?? new InlineSequence();
            }
            return copy;
        }

        var cells = new InlineSequence[expectedCount];
        int limit = Math.Min(expectedCount, row.Count);
        for (int i = 0; i < limit; i++) {
            cells[i] = row[i] ?? new InlineSequence();
        }
        if (limit < expectedCount) {
            for (int i = limit; i < expectedCount; i++) {
                cells[i] = new InlineSequence();
            }
        }
        return cells;
    }

    private IReadOnlyList<TableCell> BuildSimpleRowCells(IReadOnlyList<string> cells) {
        if (cells == null || cells.Count == 0) {
            return Array.Empty<TableCell>();
        }

        var typedCells = new TableCell[cells.Count];
        for (int i = 0; i < cells.Count; i++) {
            typedCells[i] = BuildSimpleCell(cells[i]);
        }
        return typedCells;
    }

    private static IReadOnlyList<TableCell> AssignTableCellLocations(IReadOnlyList<TableCell> cells, bool isHeader, int rowIndex) {
        for (int i = 0; i < cells.Count; i++) {
            var cell = cells[i];
            if (cell == null) {
                continue;
            }

            cell.IsHeader = isHeader;
            cell.RowIndex = isHeader ? -1 : rowIndex;
            cell.ColumnIndex = i;
        }

        return cells;
    }

    private TableCell BuildSimpleCell(string? cell) {
        if (string.IsNullOrEmpty(cell)) {
            return new TableCell();
        }

        var structuredBlocks = TryParseStructuredCellBlocks(cell);
        if (structuredBlocks != null) {
            return new TableCell(structuredBlocks);
        }

        var normalized = NormalizeBreakMarkers(cell ?? string.Empty);
        var sanitized = SanitizeInlineMarkdownInput(normalized);
        var inlines = MarkdownReader.ParseInlineText(sanitized, InlineRenderOptions, InlineRenderState);
        return new TableCell(new[] {
            new ParagraphBlock(inlines)
        });
    }

    private IReadOnlyList<IMarkdownBlock>? TryParseStructuredCellBlocks(string? cell) {
        if (string.IsNullOrEmpty(cell)) {
            return null;
        }

        var normalized = NormalizeBreakMarkers(cell ?? string.Empty);
        if (!LooksLikeStructuredMarkdownCell(normalized)) {
            return null;
        }

        var options = InlineRenderOptions == null
            ? new MarkdownReaderOptions()
            : CloneOptionsWithoutTables(InlineRenderOptions);
        var state = InlineRenderState == null
            ? new MarkdownReaderState()
            : CloneState(InlineRenderState);
        var blocks = MarkdownReader.ParseBlockFragment(normalized, options, state);
        if (blocks.Count == 0) {
            return null;
        }

        if (ContainsUnsafeRawHtmlTableCellBlocks(blocks)) {
            return null;
        }

        if (blocks.Count == 1 && blocks[0] is ParagraphBlock) {
            return null;
        }

        return blocks;
    }

    internal static bool LooksLikeStructuredMarkdownCell(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var normalized = value!;
        if (normalized.IndexOf('\n') >= 0) {
            return true;
        }

        var trimmed = normalized.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        if (trimmed.StartsWith("```", StringComparison.Ordinal)
            || trimmed.StartsWith("~~~", StringComparison.Ordinal)
            || trimmed.StartsWith(">", StringComparison.Ordinal)
            || trimmed.StartsWith("<", StringComparison.Ordinal)) {
            return true;
        }

        if (trimmed[0] == '#') {
            int run = 1;
            while (run < trimmed.Length && trimmed[run] == '#') {
                run++;
            }

            if (run <= 6 && run < trimmed.Length && char.IsWhiteSpace(trimmed[run])) {
                return true;
            }
        }

        if (trimmed.Length >= 2
            && (trimmed[0] == '-' || trimmed[0] == '*' || trimmed[0] == '+')
            && char.IsWhiteSpace(trimmed[1])) {
            return true;
        }

        int digitIndex = 0;
        while (digitIndex < trimmed.Length && char.IsDigit(trimmed[digitIndex])) {
            digitIndex++;
        }

        if (digitIndex > 0
            && digitIndex + 1 < trimmed.Length
            && (trimmed[digitIndex] == '.' || trimmed[digitIndex] == ')')
            && char.IsWhiteSpace(trimmed[digitIndex + 1])) {
            return true;
        }

        return false;
    }

    internal static bool ContainsUnsafeRawHtmlTableCellBlocks(IReadOnlyList<IMarkdownBlock> blocks) {
        for (int i = 0; i < blocks.Count; i++) {
            if (ContainsUnsafeRawHtmlTableCellBlock(blocks[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsUnsafeRawHtmlTableCellBlock(IMarkdownBlock? block) {
        if (block == null) {
            return false;
        }

        if (block is HtmlRawBlock or HtmlCommentBlock) {
            return true;
        }

        if (block is not IChildMarkdownBlockContainer container || container.ChildBlocks.Count == 0) {
            return false;
        }

        for (int i = 0; i < container.ChildBlocks.Count; i++) {
            if (ContainsUnsafeRawHtmlTableCellBlock(container.ChildBlocks[i])) {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<string> PrepareStructuredRowMarkdown(
        IReadOnlyList<TableCell>? structuredRow,
        IReadOnlyList<string>? fallbackRow,
        int expectedCount) {
        if (structuredRow == null || structuredRow.Count == 0) {
            return PrepareRowCells(fallbackRow, expectedCount);
        }

        var cells = PrepareStructuredRowCells(structuredRow, expectedCount);
        var markdown = new string[cells.Count];
        for (int i = 0; i < cells.Count; i++) {
            markdown[i] = cells[i]?.Markdown ?? string.Empty;
        }
        return markdown;
    }

    private static List<TableCell> CloneStructuredRow(IReadOnlyList<TableCell> row) {
        var cloned = new List<TableCell>(row.Count);
        for (int i = 0; i < row.Count; i++) {
            cloned.Add(CloneStructuredCell(row[i]));
        }
        return cloned;
    }

    private static TableCell CloneStructuredCell(TableCell? cell) {
        if (cell == null) {
            return new TableCell();
        }

        return new TableCell(cell.Blocks) {
            IsHeader = cell.IsHeader,
            RowIndex = cell.RowIndex,
            ColumnIndex = cell.ColumnIndex,
            SourceSpan = cell.SourceSpan,
            SyntaxChildren = cell.SyntaxChildren
        };
    }

    private static MarkdownReaderOptions CloneOptionsWithoutTables(MarkdownReaderOptions source) {
        var clone = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO);
        clone.FrontMatter = false;
        clone.Callouts = source.Callouts;
        clone.Headings = source.Headings;
        clone.FencedCode = source.FencedCode;
        clone.IndentedCodeBlocks = source.IndentedCodeBlocks;
        clone.Images = source.Images;
        clone.UnorderedLists = source.UnorderedLists;
        clone.TaskLists = source.TaskLists;
        clone.OrderedLists = source.OrderedLists;
        clone.Tables = false;
        clone.DefinitionLists = source.DefinitionLists;
        clone.TocPlaceholders = source.TocPlaceholders;
        clone.Footnotes = source.Footnotes;
        clone.PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions;
        clone.HtmlBlocks = source.HtmlBlocks;
        clone.Paragraphs = source.Paragraphs;
        clone.AutolinkUrls = source.AutolinkUrls;
        clone.AutolinkWwwUrls = source.AutolinkWwwUrls;
        clone.AutolinkWwwScheme = source.AutolinkWwwScheme;
        clone.AutolinkEmails = source.AutolinkEmails;
        clone.BackslashHardBreaks = source.BackslashHardBreaks;
        clone.InlineHtml = source.InlineHtml;
        clone.BaseUri = source.BaseUri;
        clone.DisallowScriptUrls = source.DisallowScriptUrls;
        clone.DisallowFileUrls = source.DisallowFileUrls;
        clone.AllowMailtoUrls = source.AllowMailtoUrls;
        clone.AllowDataUrls = source.AllowDataUrls;
        clone.AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls;
        clone.RestrictUrlSchemes = source.RestrictUrlSchemes;
        clone.AllowedUrlSchemes = source.AllowedUrlSchemes;
        clone.MaxInputCharacters = source.MaxInputCharacters;
        clone.InputNormalization = source.InputNormalization == null
            ? new MarkdownInputNormalizationOptions()
            : source.InputNormalization;
        clone.FencedBlockExtensions.Clear();
        for (int i = 0; i < source.FencedBlockExtensions.Count; i++) {
            clone.FencedBlockExtensions.Add(source.FencedBlockExtensions[i]);
        }

        clone.BlockParserExtensions.Clear();
        for (int i = 0; i < source.BlockParserExtensions.Count; i++) {
            clone.BlockParserExtensions.Add(source.BlockParserExtensions[i]);
        }

        clone.InlineParserExtensions.Clear();
        for (int i = 0; i < source.InlineParserExtensions.Count; i++) {
            clone.InlineParserExtensions.Add(source.InlineParserExtensions[i]);
        }

        for (int i = 0; i < source.DocumentTransforms.Count; i++) {
            if (source.DocumentTransforms[i] != null) {
                clone.DocumentTransforms.Add(source.DocumentTransforms[i]);
            }
        }

        return clone;
    }

    private static MarkdownReaderState CloneState(MarkdownReaderState state) {
        var clone = new MarkdownReaderState();
        foreach (var kvp in state.LinkRefs) {
            clone.LinkRefs[kvp.Key] = kvp.Value;
        }

        clone.SourceLineOffset = state.SourceLineOffset;
        clone.SourceTextMap = state.SourceTextMap;
        return clone;
    }

    private IReadOnlyList<IMarkdownBlock> BuildChildBlocks() {
        var blocks = new List<IMarkdownBlock>();
        var headerCells = HeaderCells;
        for (int i = 0; i < headerCells.Count; i++) {
            for (int j = 0; j < headerCells[i].Blocks.Count; j++) {
                blocks.Add(headerCells[i].Blocks[j]);
            }
        }

        var rowCells = RowCells;
        for (int rowIndex = 0; rowIndex < rowCells.Count; rowIndex++) {
            for (int cellIndex = 0; cellIndex < rowCells[rowIndex].Count; cellIndex++) {
                var cell = rowCells[rowIndex][cellIndex];
                for (int blockIndex = 0; blockIndex < cell.Blocks.Count; blockIndex++) {
                    blocks.Add(cell.Blocks[blockIndex]);
                }
            }
        }

        return blocks;
    }

    private void InvalidateRealizedCellCache() {
        _cachedHeaderCells = null;
        _cachedRowCells = null;
        _cachedCellContentSignature = null;
        _cachedUsesStructuredCells = false;
        _cachedCellColumnCount = -1;
    }

    internal int ComputeContentSignature() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + Headers.Count;
            for (int i = 0; i < Headers.Count; i++) {
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Headers[i] ?? string.Empty);
            }

            hash = (hash * 31) + Rows.Count;
            for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
                var row = Rows[rowIndex];
                hash = (hash * 31) + (row?.Count ?? -1);
                if (row == null) {
                    continue;
                }

                for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                    hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(row[cellIndex] ?? string.Empty);
                }
            }

            return hash;
        }
    }

    internal static bool TryConsumeBreakTag(string value, int index, out int consumed) {
        consumed = 0;
        int length = value.Length;
        if (index + 3 >= length) {
            return false;
        }

        if (value[index] != '<') return false;
        if (!IsSpecificLetter(value[index + 1], 'b')) return false;
        if (!IsSpecificLetter(value[index + 2], 'r')) return false;

        int position = index + 3;

        while (position < length && char.IsWhiteSpace(value[position])) {
            position++;
        }

        if (position < length && value[position] == '/') {
            position++;
            while (position < length && char.IsWhiteSpace(value[position])) {
                position++;
            }
        }

        if (position < length && value[position] == '>') {
            consumed = position - index + 1;
            return true;
        }

        return false;
    }

    internal static bool IsSpecificLetter(char value, char expected) {
        return char.ToLowerInvariant(value) == expected;
    }

    internal IReadOnlyList<MarkdownSyntaxNode> BuildSyntaxChildren(MarkdownSourceSpan? span) {
        if (!span.HasValue) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>();
        int line = span.Value.StartLine;
        int columnCount = GetEffectiveColumnCount();
        var bodyRows = RowCells;

        if (Headers.Count > 0) {
            var headerCells = HeaderCells;
            var headerChildren = BuildTableCellSyntaxChildren(
                PrepareRowCells(Headers, columnCount),
                headerCells,
                new MarkdownSourceSpan(line, line));
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableHeader,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(headerChildren) ?? new MarkdownSourceSpan(line, line),
                string.Join(" | ", Headers),
                headerChildren));
            line += 2;
        }

        for (int i = 0; i < Rows.Count; i++) {
            if (line > span.Value.EndLine) {
                break;
            }

            var rowCells = i < bodyRows.Count ? bodyRows[i] : Array.Empty<TableCell>();
            var rowChildren = BuildTableCellSyntaxChildren(
                PrepareRowCells(Rows[i], columnCount),
                rowCells,
                new MarkdownSourceSpan(line, line));
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableRow,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(rowChildren) ?? new MarkdownSourceSpan(line, line),
                string.Join(" | ", Rows[i]),
                rowChildren));
            line++;
        }

        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildTableCellSyntaxChildren(
        IReadOnlyList<string> rawCells,
        IReadOnlyList<TableCell> structuredCells,
        MarkdownSourceSpan rowSpan) {
        int cellCount = Math.Max(rawCells?.Count ?? 0, structuredCells?.Count ?? 0);
        if (cellCount == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>(cellCount);
        for (int i = 0; i < cellCount; i++) {
            string literal = rawCells != null && i < rawCells.Count
                ? rawCells[i] ?? string.Empty
                : structuredCells != null && i < structuredCells.Count
                    ? structuredCells[i]?.Markdown ?? string.Empty
                    : string.Empty;
            var cellSpan = structuredCells != null && i < structuredCells.Count
                ? structuredCells[i]?.SourceSpan ?? rowSpan
                : rowSpan;

            IReadOnlyList<MarkdownSyntaxNode> children;
            if (structuredCells != null && i < structuredCells.Count && structuredCells[i]?.SyntaxChildren != null && structuredCells[i]!.SyntaxChildren!.Count > 0) {
                children = structuredCells[i]!.SyntaxChildren!;
            } else if (structuredCells != null && i < structuredCells.Count && structuredCells[i] != null && structuredCells[i].Blocks.Count > 0) {
                var blockNodes = new List<MarkdownSyntaxNode>(structuredCells[i].Blocks.Count);
                for (int blockIndex = 0; blockIndex < structuredCells[i].Blocks.Count; blockIndex++) {
                    blockNodes.Add(MarkdownBlockSyntaxBuilder.BuildBlock(structuredCells[i].Blocks[blockIndex]));
                }
                children = blockNodes;
            } else {
                children = Array.Empty<MarkdownSyntaxNode>();
            }

            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableCell,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(children) ?? cellSpan,
                literal,
                children,
                structuredCells != null && i < structuredCells.Count ? structuredCells[i] : null));
        }

        return nodes;
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Table,
            span,
            ((IMarkdownBlock)this).RenderMarkdown(),
            BuildSyntaxChildren(span),
            this);
}
