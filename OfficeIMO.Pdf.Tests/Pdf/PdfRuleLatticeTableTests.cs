using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

/// <summary>
/// Tables read from the rules a page paints, built the way spreadsheet and report exports draw them, with no
/// structure tags to fall back on.
/// </summary>
public sealed class PdfRuleLatticeTableTests {
    private static readonly string[][] Market = {
        new[] { "ID", "Market", "Region", "Status", "Amount" },
        new[] { "1", "Alpha Futures", "EU", "Open", "10.50" },
        new[] { "2", "Beta Options", "US", "Closed", "7.25" },
        new[] { "3", "Gamma Swaps", "APAC", "Open", "99.00" },
        new[] { "4", "Delta Futures", "EU", "Hold", "1.75" },
        new[] { "5", "Epsilon Spreads", "US", "Open", "64.10" },
        new[] { "6", "Zeta Options", "APAC", "Closed", "3.30" }
    };

    [Fact]
    public void SpreadsheetBordersPaintedAsColouredBarsOverFillsReadAsOneTable() {
        // Excel-style paint: cell fills first, then borders as 0.6 pt filled bars in several colours, column
        // dividers drawn only in the header (white on the header fill), a logo image elsewhere on the page, a
        // wrapped cell and an empty one.
        string[][] expected = Market.Select(row => row.ToArray()).ToArray();
        expected[3][1] = "Gamma Swaps and Options";
        expected[4][2] = string.Empty;
        double[] x = { 72, 112, 272, 332, 392, 472 };
        double[] y = { 700, 680, 664, 648, 622, 606, 590, 574 };
        var content = new StringBuilder();
        content.Append("q 60 0 0 24 72 740 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");
        Fill(content, "0.44 0.77 0.91 rg", x[0], y[1], x[5] - x[0], y[0] - y[1]);
        for (int row = 2; row < y.Length - 1; row += 2) Fill(content, "0.85 g", x[0], y[row + 1], x[5] - x[0], y[row] - y[row + 1]);
        string[][] painted = expected.Select(row => row.ToArray()).ToArray();
        painted[3][1] = string.Empty;
        Text(content, painted, x, y, "1 g");
        content.Append("BT /F1 9 Tf 0 g 1 0 0 1 115 ").Append(Number(y[4] + 14)).Append(" Tm (Gamma Swaps and) Tj ET\n");
        content.Append("BT /F1 9 Tf 0 g 1 0 0 1 115 ").Append(Number(y[4] + 4)).Append(" Tm (Options) Tj ET\n");
        for (int row = 0; row < y.Length; row++) {
            // Each row rule is painted per colour run: black over the first columns, grey over the rest.
            Fill(content, "0 g", x[0], y[row] - 0.3, x[2] - x[0], 0.6);
            Fill(content, "0.65 g", x[2], y[row] - 0.3, x[5] - x[2], 0.6);
        }
        Fill(content, "0 g", x[0] - 0.3, y[^1], 0.6, y[0] - y[^1]);
        Fill(content, "0 g", x[5] - 0.3, y[^1], 0.6, y[0] - y[^1]);
        Fill(content, "0.65 g", x[1] - 0.3, y[^1], 0.6, y[0] - y[^1]);
        for (int column = 2; column < 5; column++) Fill(content, "1 g", x[column] - 0.3, y[1], 0.6, y[0] - y[1]);

        PdfLogicalTable table = Assert.Single(Read(content).Tables);

        Assert.Contains(table.Evidence, evidence => evidence.Code == "table-ruling-lattice");
        Assert.Equal(expected, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void BoxWithoutRowRulesKeepsEveryRowSeparate() {
        // Column dividers and a header rule, but no rules between body rows: one ruled row holds every data
        // line, so the grid must not fold them into a single row.
        double[] x = { 72, 112, 272, 332, 392, 472 };
        double[] frame = { 700, 680, 584 };
        var content = new StringBuilder();
        Text(content, Market, x, Rows(700, 20, 16, Market.Length - 1), "0 g");
        content.Append("0 G 0.5 w\n");
        foreach (double rule in frame) Line(content, x[0], rule, x[5], rule);
        foreach (double divider in x) Line(content, divider, frame[0], divider, frame[2]);

        PdfLogicalTable table = Assert.Single(Read(content).Tables);

        Assert.DoesNotContain(table.Evidence, evidence => evidence.Code == "table-ruling-lattice");
        Assert.Equal(Market, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void LabelCentredInAVerticallyMergedCellBelongsToItsTopRow() {
        string[][] rows = {
            new[] { "Group", "Item", "Qty", "Price" },
            new[] { "Hardware", "Widget", "12", "1.50" },
            new[] { "", "Sprocket", "7", "2.25" },
            new[] { "Services", "Install", "1", "90.00" }
        };
        double[] x = { 72, 172, 272, 332, 412 };
        double[] y = Rows(700, 16, 16, rows.Length - 1);
        var content = new StringBuilder();
        var cells = rows.Select(row => row.ToArray()).ToArray();
        cells[1][0] = string.Empty;
        Text(content, cells, x, y, "0 g");
        // The merged label sits halfway down the two rows it spans, just inside the lower one.
        content.Append("BT /F1 9 Tf 0 g 1 0 0 1 75 ").Append(Number(y[2] - 6)).Append(" Tm (Hardware) Tj ET\n");
        content.Append("0 G 0.5 w\n");
        for (int row = 0; row < y.Length; row++) {
            // The rule between the two merged rows skips the first column.
            Line(content, row == 2 ? x[1] : x[0], y[row], x[^1], y[row]);
        }
        foreach (double divider in x) Line(content, divider, y[0], divider, y[^1]);

        PdfLogicalTable table = Assert.Single(Read(content).Tables);

        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    private static double[] Rows(double top, double header, double body, int bodyRows) {
        var y = new double[bodyRows + 2];
        y[0] = top;
        y[1] = top - header;
        for (int row = 2; row < y.Length; row++) y[row] = y[row - 1] - body;
        return y;
    }

    private static void Text(StringBuilder content, string[][] rows, double[] x, double[] y, string headerColor) {
        for (int row = 0; row < rows.Length; row++) {
            for (int column = 0; column < rows[row].Length; column++) {
                if (rows[row][column].Length == 0) continue;
                content.Append("BT /F1 9 Tf ").Append(row == 0 ? headerColor : "0 g").Append(" 1 0 0 1 ")
                    .Append(Number(x[column] + 3)).Append(' ').Append(Number(y[row + 1] + 5))
                    .Append(" Tm (").Append(rows[row][column]).Append(") Tj ET\n");
            }
        }
    }

    private static void Fill(StringBuilder content, string color, double left, double bottom, double width, double height) =>
        content.Append(color).Append(' ').Append(Number(left)).Append(' ').Append(Number(bottom)).Append(' ')
            .Append(Number(width)).Append(' ').Append(Number(height)).Append(" re f\n");

    private static void Line(StringBuilder content, double x1, double y1, double x2, double y2) =>
        content.Append(Number(x1)).Append(' ').Append(Number(y1)).Append(" m ")
            .Append(Number(x2)).Append(' ').Append(Number(y2)).Append(" l S\n");

    private static string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private static PdfLogicalPage Read(StringBuilder content) {
        string stream = content.ToString();
        string[] objects = {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            "<< /Length " + Encoding.ASCII.GetByteCount(stream).ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + stream + "endstream",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
        };
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(Encoding.ASCII.GetByteCount(builder.ToString()));
            builder.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = Encoding.ASCII.GetByteCount(builder.ToString());
        builder.Append("xref\n0 6\n0000000000 65535 f \n");
        foreach (int offset in offsets) builder.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        builder.Append("trailer\n<< /Root 1 0 R /Size 6 >>\nstartxref\n").Append(xref.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        using PdfDocument document = PdfDocument.Load(Encoding.ASCII.GetBytes(builder.ToString()));
        return Assert.Single(document.Read().Pages);
    }
}
