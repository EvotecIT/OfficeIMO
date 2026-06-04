using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    internal void AddRow(RowBlock row) { AddBlock(row); }

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

                Guard.AbsoluteUri(kv.Value, nameof(links));
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
