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
