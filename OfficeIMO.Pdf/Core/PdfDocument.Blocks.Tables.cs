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
    /// PdfDocument.Create()
    ///     .Row(row => row
    ///         .Gap(16)
    ///         .Column(35, column => column.H2("Signals").Bullets(new[] { "Healthy", "Watch", "Needs action" }))
    ///         .Column(65, column => column.Panel("Right-side report callout.")))
    ///     .Save("report.pdf");
    /// </code>
    /// </example>
    public PdfDocument Row(System.Action<PdfRowCompose> build) {
        Guard.NotNull(build, nameof(build));
        var row = new PdfRowCompose(this);
        build(row);
        row.Commit();
        return this;
    }

    /// <summary>Adds a simple table from rows of string arrays.</summary>
    public PdfDocument Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        AddBlock(new TableBlock(rows, align, style));
        return this;
    }

    /// <summary>Adds a table from explicit cells, including optional column spans.</summary>
    public PdfDocument Table(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        AddBlock(new TableBlock(rows, align, style));
        return this;
    }

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
    public PdfDocument TableDeferred(System.Func<System.Collections.Generic.IEnumerable<string[]>> rowFactory, int batchSize = 256, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        Guard.NotNull(rowFactory, nameof(rowFactory));
        AddBlock(new DeferredTableBlock(() => ConvertDeferredTextRows(rowFactory), batchSize, align, style));
        return this;
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
    public PdfDocument TableDeferred(System.Func<System.Collections.Generic.IEnumerable<PdfTableCell[]>> rowFactory, int batchSize = 256, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        AddBlock(new DeferredTableBlock(rowFactory, batchSize, align, style));
        return this;
    }

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
    public PdfDocument KeyValueTable(System.Collections.Generic.IEnumerable<System.Collections.Generic.KeyValuePair<string, string?>> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") {
        Guard.NotNull(rows, nameof(rows));
        var keyValueRows = new System.Collections.Generic.List<PdfKeyValueRow>();
        foreach (System.Collections.Generic.KeyValuePair<string, string?> row in rows) {
            keyValueRows.Add(new PdfKeyValueRow(row.Key, row.Value));
        }

        return KeyValueTable(keyValueRows, align, style, includeHeader, keyHeader, valueHeader);
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
    public PdfDocument KeyValueTable(System.Collections.Generic.IEnumerable<PdfKeyValueRow> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") {
        AddBlock(CreateKeyValueTableBlock(rows, align, style, includeHeader, keyHeader, valueHeader));
        return this;
    }

    internal static TableBlock CreateKeyValueTableBlock(System.Collections.Generic.IEnumerable<PdfKeyValueRow> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") {
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

    internal static TableBlock CreateTableBlockWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        var tb = new TableBlock(rows, align, style);
        if (links != null) {
            foreach (var kv in links) {
                if (kv.Key.Row < 0 || kv.Key.Col < 0) {
                    throw new System.ArgumentOutOfRangeException(nameof(links), "Table link row and column indexes must be non-negative.");
                }

                if (kv.Key.Row >= tb.Rows.Count) {
                    throw new System.ArgumentOutOfRangeException(nameof(links), "Table link row index must refer to an existing table row.");
                }

                if (kv.Key.Col >= tb.Rows[kv.Key.Row].Length) {
                    throw new System.ArgumentOutOfRangeException(nameof(links), "Table link column index must refer to an existing cell in the target row.");
                }

                Guard.UriAction(kv.Value, nameof(links));
                tb.AddLink(kv.Key, kv.Value);
            }
        }

        return tb;
    }

    /// <summary>
    /// Adds a table and attaches link URIs to specific cells.
    /// </summary>
    public PdfDocument TableWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        AddBlock(CreateTableBlockWithLinks(rows, links, align, style));
        return this;
    }
}
