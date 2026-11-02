using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGeneratedOptionalBreakControlTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void GeneratedOptionalBreaksRetainDirectionalControlsAndOriginalDiagnosticOffsets(bool cff, bool named) {
        byte[] data = cff ? File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!)
            : ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'A');
        var options = named
            ? new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Mixed controls", data))
            : new PdfCreateOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, "Mixed controls");
        PdfDocument Document(string text) => PdfDocument.Create(options).Paragraph(paragraph => {
            if (named) paragraph.FontFamily("Mixed controls");
            paragraph.Text(text);
        });

        PdfDocument valid = Document("A\u200B\u200EA");
        Assert.Empty(valid.AnalyzeTextEncoding());
        Assert.Equal("A\u200EA", PdfReadDocument.Open(valid.ToBytes()).ExtractText().Trim());
        // A missing ordinary scalar stays at its authored UTF-16 offset even
        // when optional-break and retained directional controls share the run.
        const string unsupported = "A\u200B\u200EA\u200B\uFFFF";
        var diagnostics = PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
            new[] { new PdfTextRun(unsupported, fontFamily: named ? "Mixed controls" : null) },
            options, PdfStandardFont.Helvetica, "Mixed source", "Body");
        Assert.Equal(new[] { "U+200E", "U+FFFF" }, diagnostics.Select(diagnostic => diagnostic.CodePoint));
        Assert.Equal(new[] { 2, 5 }, diagnostics.Select(diagnostic => diagnostic.Index));
        Assert.Contains("verified blank glyphs", diagnostics[0].Remediation, StringComparison.Ordinal);
        Assert.Equal("missing-embedded-font-glyph", diagnostics[1].Code);
        Assert.Contains("at index 5", diagnostics[1].Message, StringComparison.Ordinal);
        Assert.All(diagnostics, diagnostic => {
            Assert.Equal("Mixed source", diagnostic.Source);
            Assert.Equal("Body.Run[0]", diagnostic.Location);
            Assert.Equal(0, diagnostic.RunIndex);
        });
        // The same complete diagnostics reach the public document preflight,
        // and unsafe text remains rejected before encoding or glyph painting.
        var documentDiagnostics = Document(unsupported).AnalyzeTextEncoding();
        Assert.Equal(diagnostics.Select(diagnostic => (diagnostic.CodePoint, diagnostic.Index, diagnostic.Code)),
            documentDiagnostics.Select(diagnostic => (diagnostic.CodePoint, diagnostic.Index, diagnostic.Code)));
        Assert.ThrowsAny<ArgumentException>(() => Document(unsupported).ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GeneratedOptionalBreaksDoNotSkipAnUnsafeDirectionalControl(bool named) {
        // This existing synthetic font intentionally draws its space. Removing
        // ZWSP is a layout rule; it cannot waive the LRM blank-glyph proof.
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(32, 'A');
        var options = named
            ? new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Unsafe mixed", data))
            : new PdfCreateOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, "Unsafe mixed");
        PdfDocument document = PdfDocument.Create(options).Paragraph(paragraph => {
            if (named) paragraph.FontFamily("Unsafe mixed");
            paragraph.Text("A\u200B\u200EA");
        });

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(document.AnalyzeTextEncoding());
        Assert.Equal("U+200E", diagnostic.CodePoint);
        Assert.Equal(2, diagnostic.Index);
        Assert.Contains("verified blank glyphs", diagnostic.Remediation, StringComparison.Ordinal);
        Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());
        // Raw font coverage still inspects the caller's raw scalar input.
        Assert.Contains(PdfTextDiagnostics.AnalyzeEmbeddedFontText("A\u200BA", data),
            raw => raw.CodePoint == "U+200B" && raw.Index == 1);
    }
}
