using System;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public sealed partial class TableBlock {
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

        columnCount = Math.Max(columnCount, GetStructuredColumnCount(StructuredHeaders));
        if (StructuredRows != null) {
            for (int rowIndex = 0; rowIndex < StructuredRows.Count; rowIndex++) {
                columnCount = Math.Max(columnCount, GetStructuredColumnCount(StructuredRows[rowIndex]));
            }
        }

        columnCount = Math.Max(columnCount, Alignments.Count);
        return columnCount;
    }

    private static int GetStructuredColumnCount(IReadOnlyList<TableCell>? row) {
        if (row == null || row.Count == 0) {
            return 0;
        }

        int columnCount = 0;
        for (int i = 0; i < row.Count; i++) {
            columnCount += Math.Max(1, row[i]?.ColumnSpan ?? 1);
        }

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
        int columnIndex = 0;
        for (int i = 0; i < cells.Count; i++) {
            var cell = cells[i];
            if (cell == null) {
                columnIndex++;
                continue;
            }

            cell.IsHeader = isHeader;
            cell.RowIndex = isHeader ? -1 : rowIndex;
            cell.ColumnIndex = columnIndex;
            columnIndex += Math.Max(1, cell.ColumnSpan);
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
}
