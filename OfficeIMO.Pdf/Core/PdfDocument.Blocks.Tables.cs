#pragma warning disable RS0026 // Text-row, cell-row and key/value overloads share the table operation names and optional layout parameters.
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    internal void AddRow(RowBlock row) { AddBlock(row); }

    /// <summary>Adds a row with percentage-based columns to the current document flow.</summary>
    /// <remarks>
    /// Rows are useful for report-style layouts where related content should sit side by side while still participating
    /// in normal pagination, margins, themes, headers, footers, and subsequent document flow. The composed row is committed
    /// once the supplied builder delegate finishes.
    /// </remarks>
    /// <param name="build">Row builder that defines column widths, layout rhythm, and column content.</param>
    /// <returns>This <see cref="PdfDocument"/> for chaining.</returns>
    /// <example>
    /// <code>
    /// PdfDocument.Create(pdf => pdf.Content(content => content
    ///     .Row(row => row
    ///         .Gap(16)
    ///         .PercentColumn(35, column => column.H2("Signals").Bullets(new[] { "Healthy", "Watch", "Needs action" }))
    ///         .PercentColumn(65, column => column.Panel("Right-side report callout.")))))
    ///     .Save("report.pdf");
    /// </code>
    /// </example>
    internal PdfDocument Row(System.Action<PdfRowBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var row = new PdfRowBuilder(this);
        build(row);
        row.Commit();
        return this;
    }

    /// <summary>Appends a table of plain text rows.</summary>
    /// <param name="rows">Rows of cell text, snapshotted during this call.</param>
    /// <param name="align">Table alignment in the document flow.</param>
    /// <param name="style">Optional table style, including column layout, alignment, header rows and the visible caption.</param>
    public PdfTable AddTable(System.Collections.Generic.IEnumerable<string[]> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) =>
        AddTableBlock(new TableBlock(rows, align, style));

    /// <summary>Appends a table of explicit cells, including spans, links, images, form fields and rich runs.</summary>
    /// <param name="rows">Rows of cells, snapshotted during this call.</param>
    /// <param name="align">Table alignment in the document flow.</param>
    /// <param name="style">Optional table style, including column layout, alignment, header rows and the visible caption.</param>
    public PdfTable AddTable(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) =>
        AddTableBlock(new TableBlock(rows, align, style));

    /// <summary>Appends a table of plain text rows with the document table style.</summary>
    /// <remarks>Kept beside the optional-parameter form so a row array is not ambiguous with <see cref="AddTable{T}(System.Collections.Generic.IEnumerable{T}, OfficeTableOptions?)"/>.</remarks>
    public PdfTable AddTable(System.Collections.Generic.IEnumerable<string[]> rows) => AddTable(rows, OfficeAlignment.Left, null);

    /// <summary>Appends a table of explicit cells with the document table style.</summary>
    /// <remarks>Kept beside the optional-parameter form so a row array is not ambiguous with <see cref="AddTable{T}(System.Collections.Generic.IEnumerable{T}, OfficeTableOptions?)"/>.</remarks>
    public PdfTable AddTable(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows) => AddTable(rows, OfficeAlignment.Left, null);

    private PdfTable AddTableBlock(TableBlock table) {
        AddBlock(table);
        return new PdfTable(table.Cells.Count, table.ColumnCount, table.Style?.Caption, isDeferred: false);
    }

    // Chaining forms used by adapters and tests that compose a document in one expression.
    internal PdfDocument Table(System.Collections.Generic.IEnumerable<string[]> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) { AddTable(rows, align, style); return this; }
    internal PdfDocument Table(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) { AddTable(rows, align, style); return this; }

    /// <summary>
    /// Adds a replayable plain-text table whose rows are enumerated and laid out in bounded batches.
    /// </summary>
    /// <remarks>
    /// The row factory is deferred until rendering and may be invoked more than once by validation and output passes,
    /// so it must return a fresh, equivalent sequence on every call. Header rows are repeated across pages and batches;
    /// footer rows, the caption, trailing spacing, and keep-with-next behavior are emitted only at the logical table edges.
    /// Automatic content-based column fitting is intentionally unsupported because it requires a global table scan;
    /// use fixed or weighted column widths instead. Row spans must remain within a batch, and every batch must resolve
    /// to the same logical column count.
    /// </remarks>
    /// <param name="rowFactory">Factory that returns a fresh sequence of table rows for each render pass.</param>
    /// <param name="batchSize">Maximum number of body rows materialized for one layout batch.</param>
    /// <param name="align">Table alignment in the document flow.</param>
    /// <param name="style">Optional table style.</param>
    public PdfTable AddDeferredTable(System.Func<System.Collections.Generic.IEnumerable<string[]>> rowFactory, int batchSize = 256, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) {
        Guard.NotNull(rowFactory, nameof(rowFactory));
        return AddDeferredTableBlock(new DeferredTableBlock(() => ConvertDeferredTextRows(rowFactory), batchSize, align, style));
    }

    /// <summary>
    /// Adds a replayable rich-cell table whose rows are enumerated and laid out in bounded batches.
    /// </summary>
    /// <remarks>
    /// The row factory is deferred until rendering and may be invoked more than once by validation and output passes,
    /// so it must return a fresh, equivalent sequence on every call. Automatic content-based column fitting is not
    /// available on this bounded path; configure fixed or weighted column widths. Row spans must remain within a batch,
    /// and every batch must resolve to the same logical column count.
    /// </remarks>
    /// <param name="rowFactory">Factory that returns a fresh sequence of rich table rows for each render pass.</param>
    /// <param name="batchSize">Maximum number of body rows materialized for one layout batch.</param>
    /// <param name="align">Table alignment in the document flow.</param>
    /// <param name="style">Optional table style.</param>
    public PdfTable AddDeferredTable(System.Func<System.Collections.Generic.IEnumerable<PdfTableCell[]>> rowFactory, int batchSize = 256, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) =>
        AddDeferredTableBlock(new DeferredTableBlock(rowFactory, batchSize, align, style));

    private PdfTable AddDeferredTableBlock(DeferredTableBlock table) {
        AddBlock(table);
        return new PdfTable(0, 0, table.Style?.Caption, isDeferred: true);
    }

    internal PdfDocument TableDeferred(System.Func<System.Collections.Generic.IEnumerable<string[]>> rowFactory, int batchSize = 256, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) { AddDeferredTable(rowFactory, batchSize, align, style); return this; }
    internal PdfDocument TableDeferred(System.Func<System.Collections.Generic.IEnumerable<PdfTableCell[]>> rowFactory, int batchSize = 256, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) { AddDeferredTable(rowFactory, batchSize, align, style); return this; }

    private static System.Collections.Generic.IEnumerable<PdfTableCell[]> ConvertDeferredTextRows(System.Func<System.Collections.Generic.IEnumerable<string[]>> rowFactory) {
        System.Collections.Generic.IEnumerable<string[]>? rows = rowFactory();
        if (rows == null) {
            throw new System.InvalidOperationException("Deferred table row factory returned null.");
        }

        foreach (string[]? row in rows) {
            if (row == null) {
                throw new System.ArgumentException("Deferred table rows cannot contain null entries.", nameof(rowFactory));
            }

            var cells = new PdfTableCell[row.Length];
            for (int cellIndex = 0; cellIndex < row.Length; cellIndex++) {
                cells[cellIndex] = new PdfTableCell(row[cellIndex]);
            }

            yield return cells;
        }
    }

    /// <summary>
    /// Adds a two-column label/value table for document metadata, invoice facts, definition lists, and similar report sections.
    /// </summary>
    /// <param name="rows">Plain text label/value rows.</param>
    /// <param name="align">Table alignment in the document flow.</param>
    /// <param name="style">Optional table style. The style is cloned before header-row settings are applied.</param>
    /// <param name="includeHeader">When true, emits a header row before the supplied values.</param>
    /// <param name="keyHeader">Header text for the label column.</param>
    /// <param name="valueHeader">Header text for the value column.</param>
    public PdfTable AddKeyValueTable(System.Collections.Generic.IEnumerable<System.Collections.Generic.KeyValuePair<string, string?>> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") {
        Guard.NotNull(rows, nameof(rows));
        var keyValueRows = new System.Collections.Generic.List<PdfKeyValueRow>();
        foreach (System.Collections.Generic.KeyValuePair<string, string?> row in rows) {
            keyValueRows.Add(new PdfKeyValueRow(row.Key, row.Value));
        }

        return AddKeyValueTable(keyValueRows, align, style, includeHeader, keyHeader, valueHeader);
    }

    /// <summary>
    /// Adds a two-column rich label/value table for document metadata, invoice facts, definition lists, and similar report sections.
    /// </summary>
    /// <param name="rows">Rich text label/value rows.</param>
    /// <param name="align">Table alignment in the document flow.</param>
    /// <param name="style">Optional table style. The style is cloned before header-row settings are applied.</param>
    /// <param name="includeHeader">When true, emits a header row before the supplied values.</param>
    /// <param name="keyHeader">Header text for the label column.</param>
    /// <param name="valueHeader">Header text for the value column.</param>
    public PdfTable AddKeyValueTable(System.Collections.Generic.IEnumerable<PdfKeyValueRow> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") =>
        AddTableBlock(CreateKeyValueTableBlock(rows, align, style, includeHeader, keyHeader, valueHeader));

    internal PdfDocument KeyValueTable(System.Collections.Generic.IEnumerable<System.Collections.Generic.KeyValuePair<string, string?>> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") { AddKeyValueTable(rows, align, style, includeHeader, keyHeader, valueHeader); return this; }
    internal PdfDocument KeyValueTable(System.Collections.Generic.IEnumerable<PdfKeyValueRow> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") { AddKeyValueTable(rows, align, style, includeHeader, keyHeader, valueHeader); return this; }

    internal static TableBlock CreateKeyValueTableBlock(System.Collections.Generic.IEnumerable<PdfKeyValueRow> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") {
        Guard.NotNull(rows, nameof(rows));
        var tableRows = new System.Collections.Generic.List<PdfTableCell[]>();
        if (includeHeader) {
            tableRows.Add(new[] {
                PdfTableCell.TextCell(keyHeader ?? string.Empty),
                PdfTableCell.TextCell(valueHeader ?? string.Empty)
            });
        }

        foreach (PdfKeyValueRow row in rows) {
            if (row == null) {
                throw new System.ArgumentException("PDF key/value table rows cannot contain null entries.", nameof(rows));
            }

            tableRows.Add(row.ToTableCells());
        }

        PdfTableStyle effectiveStyle = style?.Clone() ?? new PdfTableStyle();
        effectiveStyle.HeaderRowCount = includeHeader ? 1 : 0;
        return new TableBlock(tableRows, align, effectiveStyle);
    }
}
