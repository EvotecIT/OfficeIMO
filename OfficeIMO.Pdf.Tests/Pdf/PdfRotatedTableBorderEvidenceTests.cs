using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfRotatedTableBorderEvidenceTests {
    [Theory]
    [InlineData(90, 0)]
    [InlineData(270, 0)]
    [InlineData(90, 1)]
    [InlineData(270, 1)]
    [InlineData(90, 2)]
    [InlineData(270, 2)]
    [InlineData(90, 3)]
    [InlineData(270, 3)]
    public void NarrowTextColumnsKeepBordersAndInteriorArtworkDistinct(int rotation, int extra) {
        // Fixed PDF text metrics reproduce the narrower column evidence observed
        // with Calibri on Windows, without depending on an installed system font.
        byte[] pdf = CreateTable(rotation, extra);
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf,
            new PdfTextLayoutOptions { ForceSingleColumn = true }).Pages);
        PdfLogicalTable table = Assert.Single(page.Tables);
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(2, table.Columns.Count);
        Assert.Equal(extra == 0 ? 6 : 7, page.VectorPrimitiveCount);
        Assert.Equal(extra is 1 or 2 ? 1 : 0, page.UnrepresentedVectorPrimitiveCount);
    }

    private static byte[] CreateTable(int rotation, int extra) {
        var content = new StringBuilder();
        string[][] rows = { new[] { "Code", "Value" }, new[] { "A", "1" }, new[] { "B", "2" } };
        for (int row = 0; row < rows.Length; row++) {
            string bottom = (708.6D - row * 19.4D).ToString("0.##", CultureInfo.InvariantCulture);
            string top = (728D - row * 19.4D).ToString("0.##", CultureInfo.InvariantCulture);
            string baseline = (717.75D - row * 19.4D).ToString("0.##", CultureInfo.InvariantCulture);
            content.Append("0.8 G 0.5 w 64 ").Append(bottom).Append(" 120 19.4 re S\n")
                .Append("124 ").Append(top).Append(" m 124 ").Append(bottom).Append(" l S\n");
            for (int column = 0; column < 2; column++) {
                content.Append("BT /F1 11 Tf 90 Tz 0 g 1 0 0 1 ").Append(68 + column * 60).Append(' ')
                    .Append(baseline).Append(" Tm (").Append(rows[row][column]).Append(") Tj ET\n");
            }
        }
        if (extra == 1) content.Append("0.8 G 0.5 w 70 698 m 120 698 l S\n");
        if (extra == 2) content.Append("0.8 G 0.5 w 90 690 m 90 704 l S\n");
        if (extra == 3) content.Append("0.8 G 0.5 w 64 689.2 m 124 689.2 l S\n");
        string[] objects = {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Rotate " + rotation.ToString(CultureInfo.InvariantCulture) +
                " /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            "<< /Length " + Encoding.ASCII.GetByteCount(content.ToString()).ToString(CultureInfo.InvariantCulture) +
                " >>\nstream\n" + content + "endstream",
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
        builder.Append("trailer\n<< /Root 1 0 R /Size 6 >>\nstartxref\n")
            .Append(xref.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
