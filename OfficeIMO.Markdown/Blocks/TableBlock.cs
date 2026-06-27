using System;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Pipe table with optional header row.
/// </summary>
public sealed partial class TableBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock {
    internal const int MaxEffectiveColumnCount = 4096;
    internal const string HeaderlessSingleRowTableMarker = "<!-- OfficeIMO:RTF:HeaderlessSingleRowTable -->";

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
    /// <summary>Structured child blocks flattened from header and body cells in table order.</summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => BuildChildBlocks();
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
    /// <summary>Optional fixed column widths in points. Null entries are sized by weights or defaults.</summary>
    public List<double?> ColumnWidthPoints { get; } = new List<double?>();
    /// <summary>Optional relative column width weights. Missing columns default to 1.0 in consumers that support widths.</summary>
    public List<double> ColumnWidthWeights { get; } = new List<double>();
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
    internal MarkdownSourceSpan? AlignmentRowSourceSpan { get; private set; }
    internal bool PreserveHeaderlessSingleRowTable { get; set; }
    internal bool UseHeaderColumnCountForRendering { get; set; }
    internal bool CellsContainRenderedMarkdown { get; set; }

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

    internal void SetAlignmentRowSourceSpan(MarkdownSourceSpan sourceSpan) {
        AlignmentRowSourceSpan = sourceSpan;
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
        Func<string?, string> escapeCell = CellsContainRenderedMarkdown ? EscapeRenderedMarkdownCell : EscapeMarkdownCell;

        if (Headers.Count > 0) {
            var sb = new StringBuilder();
            var headerMarkdown = useStructuredCells
                ? PrepareStructuredRowMarkdown(StructuredHeaders, Headers, columnCount)
                : PrepareRowCells(Headers, columnCount);
            var escapedHeaders = headerMarkdown.Select(escapeCell).ToArray();
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
                var escapedRow = rowMarkdown.Select(escapeCell).ToArray();
                AppendRow(sb, escapedRow);
            }

            return sb.ToString().TrimEnd('\n');
        }

        var sbNoHeaders = new StringBuilder();
        if (PreserveHeaderlessSingleRowTable && Rows.Count == 1) {
            sbNoHeaders.Append(HeaderlessSingleRowTableMarker).Append('\n');
        }

        for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
            var rowMarkdown = useStructuredCells && StructuredRows != null && rowIndex < StructuredRows.Count
                ? PrepareStructuredRowMarkdown(StructuredRows[rowIndex], Rows[rowIndex], columnCount)
                : PrepareRowCells(Rows[rowIndex], columnCount);
            var escapedRow = rowMarkdown.Select(escapeCell).ToArray();
            AppendRow(sbNoHeaders, escapedRow);
        }
        return sbNoHeaders.ToString().TrimEnd('\n');
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        StringBuilder sb = new StringBuilder();
        sb.Append("<table>");
        AppendColumnGroupHtml(sb, GetEffectiveColumnCount());
        var headerCells = HeaderCells;
        var rowCells = RowCells;
        var headerInlines = BuildHeaderInlines();
        var rowInlines = BuildRowInlines();
        if (Headers.Count > 0) {
            sb.Append("<thead><tr>");
            int columnCount = GetEffectiveColumnCount();
            var preparedHeaders = PrepareRowCells(Headers, columnCount);
            var currentStructuredHeaders = GetCurrentStructuredHeaders();
            var preparedStructuredHeaders = PrepareStructuredRowHtmlCells(currentStructuredHeaders, columnCount, UseHeaderColumnCountForRendering)
                ?? PrepareStructuredRowCells(headerCells, columnCount);
            var preparedParsedHeaders = PrepareParsedRowCells(headerInlines, columnCount);
            int headerRenderCount = GetHtmlRenderCellCount(preparedHeaders.Count, preparedStructuredHeaders);
            for (int i = 0; i < headerRenderCount; i++) {
                var h = preparedHeaders[i];
                var style = GetAlignment(i);
                TableCell? structuredCell = preparedStructuredHeaders?[i];
                sb.Append($"<th{RenderCellAttributes(structuredCell, style)}>");
                sb.Append(RenderCellHtml(h, structuredCell, preparedParsedHeaders?[i]));
                sb.Append("</th>");
            }
            sb.Append("</tr></thead>");
        }
        if (Rows.Count > 0) {
            sb.Append("<tbody>");
            int bodyColumnCount = GetEffectiveColumnCount();
            for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
                var row = Rows[rowIndex];
                var cells = PrepareRowCells(row, bodyColumnCount);
                var currentStructuredRow = GetCurrentStructuredRow(rowIndex);
                var structuredCells = PrepareStructuredRowHtmlCells(currentStructuredRow, bodyColumnCount, UseHeaderColumnCountForRendering)
                    ?? (rowIndex < rowCells.Count ? PrepareStructuredRowCells(rowCells[rowIndex], bodyColumnCount) : null);
                var parsedCells = rowIndex < rowInlines.Count
                    ? PrepareParsedRowCells(rowInlines[rowIndex], bodyColumnCount)
                    : null;
                sb.Append("<tr>");
                int renderCellCount = GetHtmlRenderCellCount(cells.Count, structuredCells);
                for (int i = 0; i < renderCellCount; i++) {
                    var cell = cells[i];
                    var style = GetAlignment(i);
                    TableCell? structuredCell = structuredCells?[i];
                    sb.Append($"<td{RenderCellAttributes(structuredCell, style)}>");
                    sb.Append(RenderCellHtml(cell, structuredCell, parsedCells?[i]));
                    sb.Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</tbody>");
        }
        sb.Append("</table>");
        return sb.ToString();
    }

    private IReadOnlyList<TableCell>? GetCurrentStructuredHeaders() {
        return StructuredContentSignature.HasValue && StructuredContentSignature.Value == ComputeContentSignature()
            ? StructuredHeaders
            : null;
    }

    private IReadOnlyList<TableCell>? GetCurrentStructuredRow(int rowIndex) {
        return StructuredContentSignature.HasValue &&
            StructuredContentSignature.Value == ComputeContentSignature() &&
            StructuredRows != null &&
            rowIndex >= 0 &&
            rowIndex < StructuredRows.Count
            ? StructuredRows[rowIndex]
            : null;
    }

    private static int GetHtmlRenderCellCount(int preparedCount, IReadOnlyList<TableCell>? structuredCells) {
        if (structuredCells == null) {
            return preparedCount;
        }

        return Math.Min(preparedCount, structuredCells.Count);
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => ChildBlocks;

    private void AppendColumnGroupHtml(StringBuilder sb, int columnCount) {
        if (!HasColumnWidthHints(columnCount)) {
            return;
        }

        double totalWeight = 0D;
        for (int i = 0; i < columnCount; i++) {
            if (i < ColumnWidthWeights.Count && ColumnWidthWeights[i] > 0D && !double.IsNaN(ColumnWidthWeights[i]) && !double.IsInfinity(ColumnWidthWeights[i])) {
                totalWeight += ColumnWidthWeights[i];
            } else {
                totalWeight += 1D;
            }
        }

        sb.Append("<colgroup>");
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            double? fixedWidth = columnIndex < ColumnWidthPoints.Count ? ColumnWidthPoints[columnIndex] : null;
            if (fixedWidth.HasValue && fixedWidth.Value > 0D && !double.IsNaN(fixedWidth.Value) && !double.IsInfinity(fixedWidth.Value)) {
                sb.Append("<col style=\"width:");
                sb.Append(fixedWidth.Value.ToString("0.###", CultureInfo.InvariantCulture));
                sb.Append("pt\">");
                continue;
            }

            if (ColumnWidthWeights.Count == 0 || totalWeight <= 0D) {
                sb.Append("<col>");
                continue;
            }

            double weight = columnIndex < ColumnWidthWeights.Count && ColumnWidthWeights[columnIndex] > 0D
                ? ColumnWidthWeights[columnIndex]
                : 1D;
            double percentage = weight / totalWeight * 100D;
            sb.Append("<col style=\"width:");
            sb.Append(percentage.ToString("0.###", CultureInfo.InvariantCulture));
            sb.Append("%\">");
        }

        sb.Append("</colgroup>");
    }

    private bool HasColumnWidthHints(int columnCount) {
        if (columnCount <= 0) {
            return false;
        }

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            if (columnIndex < ColumnWidthPoints.Count && ColumnWidthPoints[columnIndex].HasValue) {
                return true;
            }
        }

        return ColumnWidthWeights.Count > 0;
    }

    private static string RenderCellAttributes(TableCell? cell, ColumnAlignment alignment) {
        if (cell == null) {
            return RenderStyleAttribute(alignment, null);
        }

        StringBuilder? attributes = null;
        if (cell.ColumnSpan > 1) {
            attributes ??= new StringBuilder();
            attributes.Append(" colspan=\"");
            attributes.Append(cell.ColumnSpan.ToString(CultureInfo.InvariantCulture));
            attributes.Append('"');
        }

        if (cell.RowSpan > 1) {
            attributes ??= new StringBuilder();
            attributes.Append(" rowspan=\"");
            attributes.Append(cell.RowSpan.ToString(CultureInfo.InvariantCulture));
            attributes.Append('"');
        }

        string styleAttribute = RenderStyleAttribute(alignment, cell);
        if (styleAttribute.Length > 0) {
            attributes ??= new StringBuilder();
            attributes.Append(styleAttribute);
        }

        return attributes?.ToString() ?? string.Empty;
    }

    private static string RenderStyleAttribute(ColumnAlignment alignment, TableCell? cell) {
        StringBuilder? style = null;
        ColumnAlignment effectiveAlignment = cell != null && cell.Alignment != ColumnAlignment.None ? cell.Alignment : alignment;
        if (effectiveAlignment != ColumnAlignment.None) {
            style ??= new StringBuilder();
            style.Append("text-align:");
            style.Append(effectiveAlignment switch {
                ColumnAlignment.Center => "center",
                ColumnAlignment.Right => "right",
                _ => "left"
            });
        }

        AppendStyleDeclaration(ref style, "background-color", cell?.BackgroundColor);
        AppendStyleDeclaration(ref style, "color", cell?.TextColor);
        if (cell?.Bold == true) {
            AppendStyleDeclaration(ref style, "font-weight", "bold");
        }

        if (cell?.Italic == true) {
            AppendStyleDeclaration(ref style, "font-style", "italic");
        }

        if (cell?.Underline == true || cell?.Strikethrough == true) {
            string decoration = cell.Underline && cell.Strikethrough ? "underline line-through" : cell.Underline ? "underline" : "line-through";
            AppendStyleDeclaration(ref style, "text-decoration", decoration);
        }

        return style == null || style.Length == 0 ? string.Empty : " style=\"" + style + "\"";
    }

    private static void AppendStyleDeclaration(ref StringBuilder? style, string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return;
        }

        if (style != null && style.Length > 0) {
            style.Append(';');
        }

        style ??= new StringBuilder();
        style.Append(name);
        style.Append(':');
        style.Append(System.Net.WebUtility.HtmlEncode(value!.Trim()));
    }
}
