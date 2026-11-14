using System.Globalization;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Pipe table with optional header row.
/// </summary>
public sealed partial class TableBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, IHtmlAppendableMarkdownBlock {
    internal const int MaxEffectiveColumnCount = 4096;
    internal const string HeaderlessSingleRowTableMarker = "<!-- OfficeIMO:RTF:HeaderlessSingleRowTable -->";

    private IReadOnlyList<TableCell>? _cachedHeaderCells;
    private IReadOnlyList<IReadOnlyList<TableCell>>? _cachedRowCells;
    private TableRow? _cachedHeaderRow;
    private IReadOnlyList<TableRow>? _cachedBodyRows;
    private int? _cachedCellContentSignature;
    private bool _cachedUsesStructuredCells;
    private int _cachedCellColumnCount = -1;
    private bool _structuredCellsAreOwned;

    /// <summary>Optional header cells.</summary>
    public List<string> Headers { get; } = new List<string>();
    /// <summary>Typed header cell content.</summary>
    public IReadOnlyList<TableCell> HeaderCells => GetOrBuildHeaderCells();
    /// <summary>Typed header row, or <c>null</c> when the table has no header.</summary>
    public TableRow? HeaderRow => GetOrBuildHeaderRow();
    /// <summary>Parsed inline representation of the current header cells.</summary>
    public IReadOnlyList<InlineSequence> HeaderInlines => BuildHeaderInlines();
    /// <summary>Data rows.</summary>
    public List<IReadOnlyList<string>> Rows { get; } = new List<IReadOnlyList<string>>();
    /// <summary>Typed row cell content.</summary>
    public IReadOnlyList<IReadOnlyList<TableCell>> RowCells => GetOrBuildRowCells();
    /// <summary>Typed body rows.</summary>
    public IReadOnlyList<TableRow> BodyRows => GetOrBuildBodyRows();
    /// <summary>Structured child blocks flattened from header and body cells in table order.</summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => BuildChildBlocks();
    /// <summary>Enumerates header and body rows in document order.</summary>
    public IEnumerable<TableRow> EnumerateRows() {
        var header = HeaderRow;
        if (header != null) {
            yield return header;
        }

        var rows = BodyRows;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            yield return rows[rowIndex];
        }
    }

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
    public List<MarkdownColumnAlignment> Alignments { get; } = new List<MarkdownColumnAlignment>();
    /// <summary>Optional fixed physical column widths. Null entries are sized by weights or defaults.</summary>
    public List<OfficeLength?> ColumnWidths { get; } = new List<OfficeLength?>();
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
    internal IReadOnlyList<TableAlignmentCellSource> AlignmentCellSources { get; private set; } = Array.Empty<TableAlignmentCellSource>();
    internal IReadOnlyList<TablePipeSource> PipeSources { get; private set; } = Array.Empty<TablePipeSource>();
    internal bool PreserveHeaderlessSingleRowTable { get; set; }
    internal bool UseHeaderColumnCountForRendering { get; set; }
    internal bool CellsContainRenderedMarkdown { get; set; }

    // When a table is produced by the reader, we keep the parse options/state so inline parsing in cells
    // (links/emphasis/etc) can honor URL safety settings and reference-style link definitions.
    internal MarkdownLoadOptions? InlineRenderOptions { get; set; }
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
        _structuredCellsAreOwned = false;
        InvalidateRealizedCellCache();
    }

    /// <summary>
    /// Transfers converter-owned structured cells into this table without cloning a model that
    /// was created exclusively for this table.
    /// </summary>
    internal void SetOwnedStructuredCells(
        List<TableCell>? headers,
        List<IReadOnlyList<TableCell>>? rows,
        int contentSignature) {
        StructuredHeaders = headers;
        StructuredRows = rows;
        StructuredContentSignature = contentSignature;
        _structuredCellsAreOwned = true;
        InvalidateRealizedCellCache();
    }

    internal void SetAlignmentRowSourceSpan(MarkdownSourceSpan sourceSpan) {
        AlignmentRowSourceSpan = sourceSpan;
    }

    internal void SetAlignmentCellSources(IReadOnlyList<TableAlignmentCellSource>? sources) {
        AlignmentCellSources = sources == null || sources.Count == 0
            ? Array.Empty<TableAlignmentCellSource>()
            : sources.ToArray();
    }

    internal void SetPipeSources(IReadOnlyList<TablePipeSource>? sources) {
        PipeSources = sources == null || sources.Count == 0
            ? Array.Empty<TablePipeSource>()
            : sources.ToArray();
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
        bool preserveMarkdownEscapes = useStructuredCells || CellsContainRenderedMarkdown;

        if (Headers.Count > 0) {
            var sb = new StringBuilder();
            var escapedHeaders = PrepareEscapedMarkdownRow(
                useStructuredCells ? StructuredHeaders : null,
                Headers,
                columnCount,
                preserveMarkdownEscapes);
            AppendTableAttributesToFirstCell(escapedHeaders);
            AppendRow(sb, escapedHeaders);

            var alignRow = new string[columnCount];
            for (int i = 0; i < columnCount; i++) {
                var a = GetAlignment(i);
                alignRow[i] = a switch { MarkdownColumnAlignment.Left => ":---", MarkdownColumnAlignment.Center => ":---:", MarkdownColumnAlignment.Right => "---:", _ => "---" };
            }
            AppendRow(sb, alignRow);

            for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
                var escapedRow = PrepareEscapedMarkdownRow(
                    useStructuredCells && StructuredRows != null && rowIndex < StructuredRows.Count
                        ? StructuredRows[rowIndex]
                        : null,
                    Rows[rowIndex],
                    columnCount,
                    preserveMarkdownEscapes);
                AppendRow(sb, escapedRow);
            }

            return sb.ToString().TrimEnd('\n');
        }

        var sbNoHeaders = new StringBuilder();
        if (PreserveHeaderlessSingleRowTable && Rows.Count == 1) {
            sbNoHeaders.Append(HeaderlessSingleRowTableMarker).Append('\n');
        }

        for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
            var escapedRow = PrepareEscapedMarkdownRow(
                useStructuredCells && StructuredRows != null && rowIndex < StructuredRows.Count
                    ? StructuredRows[rowIndex]
                    : null,
                Rows[rowIndex],
                columnCount,
                preserveMarkdownEscapes);
            if (rowIndex == 0) {
                AppendTableAttributesToFirstCell(escapedRow);
            }
            AppendRow(sbNoHeaders, escapedRow);
        }
        return sbNoHeaders.ToString().TrimEnd('\n');
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }

    void IHtmlAppendableMarkdownBlock.AppendHtml(StringBuilder builder) => AppendHtml(builder);

    private void AppendHtml(StringBuilder sb) {
        bool[] numericColumns = ResolveNumericColumns();
        sb.Append("<table").Append(MarkdownHtmlAttributes.Render(Attributes, null)).Append(">");
        int columnCount = GetEffectiveColumnCount();
        AppendColumnGroupHtml(sb, columnCount);
        // The cell text is hashed once for the whole render.
        int contentSignature = ComputeContentSignature();
        EnsureRealizedCells(contentSignature, columnCount);
        var headerCells = _cachedHeaderCells ?? Array.Empty<TableCell>();
        var rowCells = _cachedRowCells ?? Array.Empty<IReadOnlyList<TableCell>>();
        bool structuredCellsAreCurrent = StructuredContentSignature.HasValue && StructuredContentSignature.Value == contentSignature;
        // Parsed cell inlines are a fallback for a row without typed cells, so they are built only then.
        IReadOnlyList<IReadOnlyList<InlineSequence>>? rowInlines = null;
        if (Headers.Count > 0) {
            sb.Append("<thead><tr>");
            var currentStructuredHeaders = structuredCellsAreCurrent ? StructuredHeaders : null;
            var htmlHeaders = GetHtmlRowCells(currentStructuredHeaders, columnCount, UseHeaderColumnCountForRendering)
                ?? GetHtmlRowCells(headerCells, columnCount);
            int headerRenderCount = Math.Min(GetPreparedCellCount(Headers, columnCount), htmlHeaders.Count);
            for (int i = 0; i < headerRenderCount; i++) {
                AppendCellHtml(sb, "th", htmlHeaders[i], GetHtmlAlignment(i, numericColumns));
            }
            sb.Append("</tr></thead>");
        }
        if (Rows.Count > 0) {
            sb.Append("<tbody>");
            for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
                var row = Rows[rowIndex];
                var currentStructuredRow = structuredCellsAreCurrent && StructuredRows != null && rowIndex < StructuredRows.Count
                    ? StructuredRows[rowIndex]
                    : null;
                var structuredCells = GetHtmlRowCells(currentStructuredRow, columnCount, UseHeaderColumnCountForRendering)
                    ?? (rowIndex < rowCells.Count ? GetHtmlRowCells(rowCells[rowIndex], columnCount) : null);
                sb.Append("<tr>");
                if (structuredCells != null) {
                    int renderCellCount = Math.Min(GetPreparedCellCount(row, columnCount), structuredCells.Count);
                    for (int i = 0; i < renderCellCount; i++) {
                        AppendCellHtml(sb, "td", structuredCells[i], GetHtmlAlignment(i, numericColumns));
                    }
                } else {
                    var cells = PrepareRowCells(row, columnCount);
                    rowInlines ??= BuildRowInlines(contentSignature, columnCount);
                    var parsedCells = rowIndex < rowInlines.Count
                        ? PrepareParsedRowCells(rowInlines[rowIndex], columnCount)
                        : null;
                    for (int i = 0; i < cells.Count; i++) {
                        sb.Append("<td").Append(RenderCellAttributes(null, GetHtmlAlignment(i, numericColumns))).Append('>');
                        sb.Append(RenderCellHtml(cells[i], null, parsedCells?[i]));
                        sb.Append("</td>");
                    }
                }
                sb.Append("</tr>");
            }
            sb.Append("</tbody>");
        }
        sb.Append("</table>");
    }

    private static void AppendCellHtml(StringBuilder sb, string tag, TableCell cell, MarkdownColumnAlignment alignment) {
        sb.Append('<').Append(tag).Append(RenderCellAttributes(cell, alignment)).Append('>');
        cell.AppendHtml(sb);
        sb.Append("</").Append(tag).Append('>');
    }

    // Number of cells PrepareRowCells would give the row.
    private static int GetPreparedCellCount(IReadOnlyList<string>? row, int columnCount) =>
        columnCount > 0 ? columnCount : row?.Count ?? 0;

    // The cells one HTML row shows: the row padded or cut to
    // the column count, with missing cells empty. Rendering only reads the cells, so the row's own cells
    // are used instead of copies.
    private static IReadOnlyList<TableCell> GetHtmlRowCells(IReadOnlyList<TableCell>? row, int columnCount) {
        int rowCount = row?.Count ?? 0;
        int count = columnCount > 0 ? columnCount : rowCount;
        if (count == 0) {
            return Array.Empty<TableCell>();
        }

        if (rowCount == count && !ContainsNullCell(row!)) {
            return row!;
        }

        var cells = new TableCell[count];
        for (int i = 0; i < count; i++) {
            cells[i] = i < rowCount ? row![i] ?? new TableCell() : new TableCell();
        }

        return cells;
    }

    // The cells one HTML row shows without padding: spans past the last
    // column are cut. Only a cell whose span is cut is copied.
    private static IReadOnlyList<TableCell>? GetHtmlRowCells(IReadOnlyList<TableCell>? row, int maxLogicalColumnCount, bool padToLogicalColumnCount) {
        if (row == null) {
            return null;
        }

        if (padToLogicalColumnCount) {
            return GetHtmlRowCells(row, maxLogicalColumnCount);
        }

        if (row.Count == 0 || maxLogicalColumnCount <= 0) {
            return Array.Empty<TableCell>();
        }

        var cells = new List<TableCell>(Math.Min(row.Count, maxLogicalColumnCount));
        int remainingColumns = maxLogicalColumnCount;
        for (int i = 0; i < row.Count && remainingColumns > 0; i++) {
            TableCell cell = row[i] ?? new TableCell();
            int columnSpan = Math.Max(1, cell.ColumnSpan);
            if (columnSpan > remainingColumns) {
                cell = CloneStructuredCell(cell);
                cell.ColumnSpan = remainingColumns;
                columnSpan = remainingColumns;
            }

            cells.Add(cell);
            remainingColumns -= columnSpan;
        }

        return cells;
    }

    private static bool ContainsNullCell(IReadOnlyList<TableCell> row) {
        for (int i = 0; i < row.Count; i++) {
            if (row[i] == null) {
                return true;
            }
        }

        return false;
    }

    private void AppendTableAttributesToFirstCell(IList<string> cells) {
        if (Attributes.IsEmpty || cells == null || cells.Count == 0) {
            return;
        }

        cells[0] = (cells[0] ?? string.Empty) + EscapeRenderedMarkdownCell(MarkdownAttributeBlockRenderer.RenderTrailing(Attributes));
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
            double? fixedWidth = columnIndex < ColumnWidths.Count ? ColumnWidths[columnIndex]?.ToPoints() : null;
            if (fixedWidth.HasValue && fixedWidth.Value > 0D) {
                sb.Append("<col style=\"width:");
                sb.Append(fixedWidth.Value.ToString("0.###", CultureInfo.InvariantCulture));
                sb.Append("pt\">");
                continue;
            }

            if (ColumnWidthWeights.Count == 0 || totalWeight <= 0D) {
                sb.Append("<col>");
                continue;
            }

            double weight = columnIndex < ColumnWidthWeights.Count && ColumnWidthWeights[columnIndex] > 0D &&
                !double.IsNaN(ColumnWidthWeights[columnIndex]) && !double.IsInfinity(ColumnWidthWeights[columnIndex])
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
            if (columnIndex < ColumnWidths.Count && ColumnWidths[columnIndex].HasValue) {
                return true;
            }
        }

        return ColumnWidthWeights.Count > 0;
    }

    private MarkdownColumnAlignment GetHtmlAlignment(int index, bool[] numericColumns) {
        MarkdownColumnAlignment authored = GetAlignment(index);
        return authored == MarkdownColumnAlignment.None && index < numericColumns.Length && numericColumns[index]
            ? MarkdownColumnAlignment.Right
            : authored;
    }

    // Styled HTML right-aligns columns of numbers that the source left unaligned, as the visual design asks.
    // Plain and GitHub-compatible output keep the source alignment exactly.
    private bool[] ResolveNumericColumns() {
        var options = HtmlRenderContext.Options;
        if (options == null || HtmlRenderer.GetEffectiveStyle(options) == MarkdownHtmlStyle.Plain || options.GitHubTaskListHtml || Rows.Count == 0) {
            return Array.Empty<bool>();
        }

        int columnCount = GetEffectiveColumnCount();
        var numeric = new bool[columnCount];
        for (int column = 0; column < columnCount; column++) {
            bool any = false;
            bool all = true;
            foreach (IReadOnlyList<string> row in Rows) {
                string value = column < row.Count ? row[column] ?? string.Empty : string.Empty;
                if (value.Trim().Length == 0) continue;
                any = true;
                if (!global::OfficeIMO.Internal.OfficeThemeCssTokens.IsNumericText(value.Replace("\\", string.Empty).Replace("*", string.Empty))) {
                    all = false;
                    break;
                }
            }
            numeric[column] = any && all;
        }

        return numeric;
    }

    private static string RenderCellAttributes(TableCell? cell, MarkdownColumnAlignment alignment) {
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

    private static string RenderStyleAttribute(MarkdownColumnAlignment alignment, TableCell? cell) {
        StringBuilder? style = null;
        MarkdownColumnAlignment effectiveAlignment = cell != null && cell.Alignment != MarkdownColumnAlignment.None ? cell.Alignment : alignment;
        if (effectiveAlignment != MarkdownColumnAlignment.None) {
            style ??= new StringBuilder();
            style.Append("text-align:");
            style.Append(effectiveAlignment switch {
                MarkdownColumnAlignment.Center => "center",
                MarkdownColumnAlignment.Right => "right",
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
        style.Append(HtmlTextEncoder.Encode(value!.Trim(), HtmlRenderContext.Options));
    }
}
