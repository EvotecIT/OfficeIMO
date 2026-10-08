using System.Globalization;
using OfficeIMO.Pdf;

/// <summary>Independent reader qualification of each PDF table cell, its column and ordered row.</summary>
internal static class DataTablesPdfComparisonVerifier {
    internal static object Verify(string path, int rows, int columns, bool unique) {
        PdfReadDocument document = PdfReadDocument.Open(path);
        Require(!document.RepairReport.HasRepairs, "PDF required reader repairs.");
        int row = 0, headings = 0;
        const double pageWidth = 1190.551, pageHeight = 841.89;
        double width = (pageWidth - 72) / columns;
        string[]? current = null;
        void CompleteRow() {
            if (current is null) return;
            Require(row < rows, "PDF exceeds the declared rows.");
            for (int column = 0; column < columns; column++)
                Require(current[column] == Expected(row, column, unique), $"PDF cell {row}/{column} differs: {current[column]} != {Expected(row, column, unique)}.");
            row++; current = null;
        }
        foreach (PdfReadPage page in document.Pages) {
            var size = page.GetPageSize();
            Require(Math.Abs(size.Width - pageWidth) < .02 && Math.Abs(size.Height - pageHeight) < .02, "PDF page size differs.");
            bool headingSeen = false;
            foreach (var line in page.GetTextSpans().GroupBy(span => Math.Round(span.Y, 2)).OrderByDescending(group => group.Key)) {
                string[] values = new string[columns];
                foreach (PdfTextSpan span in line.OrderBy(span => span.X)) {
                    Require(span.IsVisible && Math.Abs(span.FontSize - 6) < .01, "PDF table text style differs.");
                    Require(span.X >= 35.9 && span.X + span.Advance <= pageWidth - 35.8, "PDF text exceeds the printable page width.");
                    int column = (int)Math.Floor((span.X - 36 + .01) / width);
                    Require(column >= 0 && column < columns, "PDF text is outside a declared column.");
                    Require(span.X + span.Advance <= 36 + (column + 1) * width - 3.5, "PDF text exceeds its cell's printable width.");
                    values[column] += span.Text;
                }
                bool heading = values[0] == "Column1";
                if (heading) {
                    Require(!headingSeen, "PDF duplicated its repeated heading.");
                    headingSeen = true; headings++;
                    for (int column = 0; column < columns; column++) Require(values[column] == "Column" + (column + 1), "PDF heading differs.");
                } else {
                    Require(headingSeen, "PDF data precedes its heading.");
                    // Column zero contains a short integer and starts each data row. Other
                    // cells may wrap, including full-precision numbers, across later baselines/pages.
                    if (values[0] is not null) { CompleteRow(); current = new string[columns]; }
                    Require(current is not null, "PDF continuation has no data row.");
                    for (int column = 0; column < columns; column++) current![column] += values[column];
                }
            }
            Require(headingSeen, "PDF page lost its repeated heading.");
        }
        CompleteRow();
        Require(row == rows && headings == document.Pages.Count, "PDF final dimensions differ.");
        return new { rows, columns, cells = (long)(rows + headings) * columns, allValues = true, pages = headings, repairs = 0,
            contract = "ordered values and column positions; repeated headings; Carlito 6-point text; A3 landscape" };
    }
    private static string Expected(int row, int column, bool unique) => column % 3 == 0 ? ((long)row * 10 + column).ToString(CultureInfo.InvariantCulture)
        : column % 3 == 1 ? $"Łódź{(unique ? row : row % 100)}-{column}" : (row % 10000 + column / 100d).ToString("R", CultureInfo.InvariantCulture);
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidDataException(message); }
}
