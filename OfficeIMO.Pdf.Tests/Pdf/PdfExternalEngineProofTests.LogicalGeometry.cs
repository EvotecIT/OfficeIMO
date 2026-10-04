#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalEngineProofTests {
    [Theory]
    [InlineData("moons", false, 0)]
    [InlineData("moons", true, 0)]
    [InlineData("Résumé 😀", false, 0)]
    [InlineData("moons", false, 1)]
    public void BoundedLogicalSpansPreserveAdjacentPunctuationAndAuthoredSpace(string word, bool textPaint, int embeddedFont) {
        var options = new PdfOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers();
        if (embeddedFont != 0) {
            byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ');
            options.EmbedStandardFont(PdfStandardFont.Helvetica, data);
        }
        double height = 12D;
        static OfficeShape FilledRectangle(double width, double height) {
            OfficeShape shape = OfficeShape.Rectangle(width, height);
            shape.FillColor = OfficeColor.Black;
            return shape;
        }
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas => {
            canvas.ActualText(word, 10D, 28D, 50D, height, paint => {
                if (textPaint) paint.Text("PaintOnly", 10D, 28D, 50D, 12D);
                else paint.Shape(FilledRectangle(50D, height), 10D, 28D);
            });
            canvas.ActualText(".", 60D, 28D, 5D, height, paint => paint.Shape(FilledRectangle(5D, height), 60D, 28D));
            canvas.ActualText(" next", 65D, 28D, 40D, height, paint => paint.Shape(FilledRectangle(40D, height), 65D, 28D));
        })), options).ToBytes();
        string expected = word + ". next";
        var read = PdfReadDocument.Open(pdf);
        PdfTextSpan first = Assert.Single(read.Pages[0].GetTextSpans(), span => span.Text == word);
        Assert.Equal(50D, first.Advance, 2);
        Assert.False(first.CanRestamp);
        Assert.Equal(expected, read.ExtractText().Trim());
        PdfExternalValidator validator = PdfExternalValidator.PdfText();
        if (!validator.IsAvailable) {
            Assert.NotEqual("1", Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_PDF_TEXT_VALIDATOR"));
            return;
        }
        PdfExternalProcessResult result = validator.Run(pdf, "bounded-logical-text.pdf");
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            string name = "bounded-logical-font" + embeddedFont + "-" + (textPaint ? "text-" : "shape-") +
                Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(word))).Substring(0, 8).ToLowerInvariant();
            File.WriteAllBytes(Path.Combine(output, name + ".pdf"), pdf);
            File.WriteAllText(Path.Combine(output, name + ".json"), JsonSerializer.Serialize(new {
                ExpectedText = expected, ActualText = result.Output.Trim(), result.ExitCode,
                result.ValidatorName, validator.ExecutablePath,
                PdfSha256 = Convert.ToHexString(SHA256.HashData(pdf)).ToLowerInvariant(),
                Passed = result.ExitCode == 0 && string.Equals(expected, result.Output.Trim(), StringComparison.Ordinal)
            }, new JsonSerializerOptions { WriteIndented = true }));
        }
        Assert.True(result.ExitCode == 0, result.GetDiagnosticText());
        Assert.Equal(expected, result.Output.Trim());
    }

    [Fact]
    public void BoundedLogicalSpanCannotBeEditedSeparatelyFromItsPaint() {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas =>
            canvas.ActualText("logical", 10D, 28D, 50D, 12D,
                paint => paint.Text("PaintOnly", 10D, 28D, 50D, 12D))))).ToBytes();
        var document = PdfDocument.Load(pdf);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans());
        Assert.Throws<NotSupportedException>(() => document.Text.Replace(
            new PdfPageRegion(1, span.X, span.Y, span.Advance, span.FontSize), "updated",
            new PdfTextEditOptions { AllowTextRenderingMode3 = true }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BoundedLogicalSpanRetainsFormulaAndLinkedTranslucentPaint(bool replacementOutsideEffect) {
        OfficeShape shape = OfficeShape.Rectangle(50D, 12D);
        shape.FillColor = OfficeColor.Blue;
        void Paint(PdfPageCanvas paint) => paint
            .Shape(shape, 10D, 28D, linkUri: "https://example.test/formula", linkContents: "logical")
            .Text("PaintOnly", 10D, 28D, 50D, 12D, 9D);
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas =>
            canvas.Structure(PdfCanvasStructureRole.Formula, formula => {
                if (replacementOutsideEffect) formula.ActualText("logical", 10D, 28D, 50D, 12D,
                    logical => logical.Effect(OfficeTransform.Translate(20D, 10D), .5D, Paint));
                else formula.Effect(OfficeTransform.Translate(20D, 10D), .5D,
                    effect => effect.ActualText("logical", 10D, 28D, 50D, 12D, Paint));
            },
                new PdfCanvasStructureOptions { AlternativeText = "A logical expression" }))),
            new PdfOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers()).ToBytes();
        var read = PdfReadDocument.Open(pdf);
        Assert.Equal("logical", read.ExtractText().Trim());
        PdfTextSpan span = Assert.Single(read.Pages[0].GetTextSpans());
        Assert.Equal(50D, span.Advance, 2);
        Assert.Equal(replacementOutsideEffect ? 10D : 30D, span.X, 2);
        Assert.Equal("https://example.test/formula", Assert.Single(read.Pages[0].GetLinkAnnotations()).Uri);
        Assert.Contains(read.TaggedContent!.StructureElements, element => element.StructureType == "Formula");
        PdfExternalValidator validator = PdfExternalValidator.PdfText();
        if (!validator.IsAvailable) {
            Assert.NotEqual("1", Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_PDF_TEXT_VALIDATOR"));
            return;
        }
        PdfExternalProcessResult result = validator.Run(pdf, "bounded-formula.pdf");
        Assert.True(result.ExitCode == 0, result.GetDiagnosticText());
        Assert.Equal("logical", result.Output.Trim());
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) File.WriteAllBytes(Path.Combine(output,
            replacementOutsideEffect ? "bounded-formula-outer.pdf" : "bounded-formula.pdf"), pdf);
    }
}
#endif
