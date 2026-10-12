using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSearchableTextRunTests {
    [Fact]
    public void OcrTextIsSelectableOnRequest() {
        // An OCR layer: rendering mode 3 text over where a scan's words appear.
        byte[] pdf = Encoding.ASCII.GetBytes(BuildPdf("BT /F1 12 Tf 3 Tr 50 700 Td (scanned words) Tj ET\n"));

        Assert.Empty(PdfPageInteractionMap.Create(pdf, 1).TextRegions);
        PdfPageInteractionRegion[] regions = PdfPageInteractionMap.Create(pdf, 1,
            new PdfPageInteractionOptions { IncludeTextRenderingMode3 = true }).TextRegions.ToArray();
        Assert.Equal("scanned words", string.Concat(regions.Select(region => region.Text)));
        Assert.InRange(regions[0].Quad.Left, 49D, 51D);
    }

    private static string BuildPdf(string content) {
        var builder = new StringBuilder("%PDF-1.7\n");
        var objects = new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            "<< /Length " + content.Length + " >>\nstream\n" + content + "endstream",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
        };
        var offsets = new int[objects.Length];
        for (int index = 0; index < objects.Length; index++) {
            offsets[index] = builder.Length;
            builder.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = builder.Length;
        builder.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) builder.Append(offset.ToString("D10")).Append(" 00000 n \n");
        return builder.Append("trailer\n<< /Size ").Append(objects.Length + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n").ToString();
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public void TransformedClustersRetainIndividualAdvancesAndOffsets(int fontKind, bool gap) {
        var options = new PdfOptions { CompressContentStreams = false };
        if (fontKind != 0) {
            string? path = fontKind == 1 ? PdfComplianceTestFonts.FindBundledTrueTypeFont() : PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
            Assert.NotNull(path);
            options.EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(path!));
        }
        string[] text = { "A", "😀", "fi", "A" };
        double[] widths = { 10, 18, 22, 7 };
        var expected = new PdfSelectionQuad[text.Length];
        double x = 30;
        for (int i = 0; i < text.Length; i++) {
            if (gap && i == 2) x += 12;
            expected[i] = new PdfSelectionQuad(Point(x, 40), Point(x + widths[i], 40), Point(x + widths[i], 60), Point(x, 60));
            x += widths[i];
        }
        byte[] bytes = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas => {
            for (int i = 0; i < text.Length; i++) canvas.SearchableText(text[i], expected[i]);
        })), options).ToBytes();
        byte[] raw = Encoding.GetEncoding(28591).GetBytes(Encoding.GetEncoding(28591).GetString(bytes).Replace("/ActualText", "/UnusedText"));
        var regions = PdfPageInteractionMap.Create(raw, 1, new PdfPageInteractionOptions { IncludeInvisibleText = true }).TextRegions.ToArray();
        Assert.Equal("A😀fiA", string.Concat(regions.Select(r => r.Text)));
        double offsetX = regions[0].Quad.BottomLeft.X - expected[0].BottomLeft.X;
        double offsetY = regions[0].Quad.BottomLeft.Y - expected[0].BottomLeft.Y;
        int offset = 0;
        for (int i = 0; i < text.Length; i++) {
            int count = text[i] == "fi" ? 2 : 1;
            Assert.InRange(Math.Abs(regions[offset].Quad.BottomLeft.X - expected[i].BottomLeft.X - offsetX), 0, .03);
            Assert.InRange(Math.Abs(regions[offset].Quad.BottomLeft.Y - expected[i].BottomLeft.Y - offsetY), 0, .03);
            Assert.True(Math.Abs(regions[offset + count - 1].Quad.BottomRight.X - expected[i].BottomRight.X - offsetX) < .03, $"Cluster {i}: {regions[offset + count - 1].Quad.BottomRight.X} versus {expected[i].BottomRight.X + offsetX}; regions {string.Join("; ", regions.Select(r => $"{r.Text}:{r.Quad.BottomLeft.X}-{r.Quad.BottomRight.X}"))}");
            offset += count;
        }
        Assert.Empty(PdfPageInteractionMap.Create(bytes, 1).TextRegions);
    }

    [Theory]
    [InlineData(.2214321D)]
    [InlineData(.0044321D)]
    public void ContiguousClustersPreserveFractionalHeightAndTotalAdvance(double height) {
        var options = new PdfOptions { CompressContentStreams = false };
        string? fontPath = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(fontPath);
        options.EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!));
        byte[] bytes = PdfDocument.Create(options).Canvas(canvas => {
            canvas.SearchableText("A", new PdfSelectionQuad(new(10, 20), new(35, 20), new(35, 20 + height), new(10, 20 + height)));
            canvas.SearchableText("B", new PdfSelectionQuad(new(35, 20), new(60, 20), new(60, 20 + height), new(35, 20 + height)));
        }).ToBytes();
        // Read glyph advances independently of the enclosing ActualText span.
        byte[] raw = Encoding.GetEncoding(28591).GetBytes(Encoding.GetEncoding(28591).GetString(bytes).Replace("/ActualText", "/UnusedText"));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(raw);
        var letters = independent.GetPage(1).Letters;
        Assert.Equal(2, letters.Count);
        Assert.Equal(10D, letters[0].StartBaseLine.X, 3);
        Assert.InRange(Math.Abs(35D - letters[0].EndBaseLine.X), 0D, .002D);
        Assert.InRange(Math.Abs(60D - letters[1].EndBaseLine.X), 0D, .002D);
    }

    private static PdfSelectionPoint Point(double x, double y) => new(x + y * .3, y + x * .2);
}
