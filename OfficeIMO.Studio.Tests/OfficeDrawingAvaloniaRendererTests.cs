using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class OfficeDrawingAvaloniaRendererTests {
    [Fact]
    public void SubstitutedPdfGlyphsUseTheRasterPreviewPath() {
        const string content = "BT /F1 20 Tf 20 80 Td (A) Tj ET";
        const string cmap = "begincmap\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n1 beginbfchar\n<41> <0051>\nendbfchar\nendcmap";
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>", "endobj",
            "4 0 obj", $"<< /Length {content.Length} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [65 /B] >> /ToUnicode 6 0 R >>", "endobj",
            "6 0 obj", $"<< /Length {cmap.Length} >>", "stream", cmap, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }) + "\n");
        OfficeDrawing drawing = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();

        Assert.Equal("Q", Assert.Single(drawing.Elements.OfType<OfficeDrawingText>()).Text);
        Assert.Contains(OfficeDrawingAvaloniaRenderer.AnalyzeRasterFallback(drawing),
            reason => reason.Contains("painted PDF glyphs", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(60D, 100D)]
    [InlineData(120D, 80D)]
    public void PositionedRunFitsItsSourceAdvanceInBothDirections(double measuredWidth, double advance) {
        (double offsetX, double scaleX) = OfficeDrawingAvaloniaRenderer.FitPositionedSingleLine(measuredWidth, advance);

        Assert.Equal(0D, offsetX);
        Assert.Equal(advance, measuredWidth * scaleX, 6);
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Left)]
    [InlineData(OfficeTextAlignment.Justify)]
    [InlineData(OfficeTextAlignment.Center)]
    [InlineData(OfficeTextAlignment.Right)]
    public void WiderSubstituteRunIsCompressedIntoItsBox(OfficeTextAlignment alignment) {
        (double offsetX, double scaleX) = OfficeDrawingAvaloniaRenderer.FitSingleLine(120D, 80D, alignment);

        Assert.Equal(0D, offsetX, 6);
        Assert.Equal(80D, 120D * scaleX, 6);
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Left, 0D)]
    [InlineData(OfficeTextAlignment.Center, 20D)]
    [InlineData(OfficeTextAlignment.Right, 40D)]
    public void NarrowerRunKeepsItsWidthAndHonorsAlignment(OfficeTextAlignment alignment, double expectedOffset) {
        (double offsetX, double scaleX) = OfficeDrawingAvaloniaRenderer.FitSingleLine(60D, 100D, alignment);

        Assert.Equal(1D, scaleX);
        Assert.Equal(expectedOffset, offsetX, 6);
    }

    [Theory]
    [InlineData(0D, 50D)]
    [InlineData(50D, 0D)]
    [InlineData(double.NaN, 50D)]
    [InlineData(50D, double.PositiveInfinity)]
    public void DegenerateMeasurementsLeaveTheRunUntransformed(double measuredWidth, double boxWidth) {
        Assert.Equal((0D, 1D), OfficeDrawingAvaloniaRenderer.FitSingleLine(measuredWidth, boxWidth, OfficeTextAlignment.Right));
    }
}
