using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfScalarCoverageProofTests {
    [Fact]
    public void RepeatedSuccessfulDiagnosticsLeaveDocumentSubsetsIndependent() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, 'A', 'B');
        PdfTrueTypeFontProgram root = PdfTrueTypeFontProgram.Parse(data, "Coverage");
        PdfTrueTypeFontProgram first = root.ForkForDocument();
        Assert.Empty(PdfTextDiagnostics.AnalyzeEmbeddedFontText("AB", first, "first"));
        Array.Clear(data, 0, data.Length);

        System.Threading.Tasks.Parallel.For(0, 8, index => {
            PdfTrueTypeFontProgram fork = root.ForkForDocument();
            Assert.Empty(PdfTextDiagnostics.AnalyzeEmbeddedFontText("AB", fork, "fork " + index));
            Assert.Empty(fork.GetUsedGlyphIds());
            fork.ShapeText("A", PdfTextShapingOptions.ForRendering("Coverage"));
            Assert.True(fork.TryGetGlyphId('A', out int glyph));
            Assert.Equal(glyph, Assert.Single(fork.GetUsedGlyphIds()));
        });
        Assert.Empty(first.GetUsedGlyphIds());
        Assert.Empty(root.GetUsedGlyphIds());
    }

    [Theory]
    [InlineData("AB", "U+0042", false)]
    [InlineData("A\u0001", "U+0001", true)]
    [InlineData("A\u007F", "U+007F", true)]
    public void RepeatedRejectedTextKeepsEachCallersWarningDetails(string text, string codePoint, bool control) {
        PdfTrueTypeFontProgram font = PdfTrueTypeFontProgram.Parse(
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, 'A'), "Coverage");
        Assert.Empty(PdfTextDiagnostics.AnalyzeEmbeddedFontText("A", font, "success"));

        foreach (string source in new[] { "first", "second" }) {
            var diagnostic = Assert.Single(PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, font, source, source + " font"));
            Assert.Equal(1, diagnostic.Index);
            Assert.Equal(source, diagnostic.Source);
            Assert.Equal(codePoint, diagnostic.CodePoint);
            Assert.Equal(control, diagnostic.IsControlCharacter);
            if (!control) Assert.Contains(source + " font", diagnostic.Message);
        }
        Assert.Empty(font.GetUsedGlyphIds());
    }

    [Fact]
    public void SuccessfulCoverageDoesNotTransferBetweenDifferentFontPrograms() {
        PdfTrueTypeFontProgram complete = PdfTrueTypeFontProgram.Parse(
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, 'A', 'B'), "Same name");
        PdfTrueTypeFontProgram incomplete = PdfTrueTypeFontProgram.Parse(
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, 'A'), "Same name");
        Assert.Empty(PdfTextDiagnostics.AnalyzeEmbeddedFontText("AB", complete));

        var diagnostic = Assert.Single(PdfTextDiagnostics.AnalyzeEmbeddedFontText("AB", incomplete, "other font"));

        Assert.Equal("U+0042", diagnostic.CodePoint);
        Assert.Equal("other font", diagnostic.Source);
        Assert.Equal(1, diagnostic.Index);
    }

    [Fact]
    public void LongCoverageAndUnicodeInputRetainTheFullDiagnosticPath() {
        PdfTrueTypeFontProgram font = PdfTrueTypeFontProgram.Parse(
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, 'A', 'B', 0x1F600), "Coverage");
        string covered = new string('A', 600) + "\U0001F600";

        for (int index = 0; index < 2; index++) {
            Assert.Empty(PdfTextDiagnostics.AnalyzeEmbeddedFontText(covered, font));
            var diagnostic = Assert.Single(PdfTextDiagnostics.AnalyzeEmbeddedFontText(covered + "Z", font, "long"));
            Assert.Equal(602, diagnostic.Index);
            Assert.Equal("U+005A", diagnostic.CodePoint);
            Assert.Equal("long", diagnostic.Source);
        }
        Assert.Empty(font.GetUsedGlyphIds());
    }
}
