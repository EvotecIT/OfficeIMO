using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSearchableTextRunTests {
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

    private static PdfSelectionPoint Point(double x, double y) => new(x + y * .3, y + x * .2);
}
