using System;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Pipe table with optional header row.
/// </summary>
public sealed partial class TableBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock, IChildMarkdownBlockContainer {
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
}
