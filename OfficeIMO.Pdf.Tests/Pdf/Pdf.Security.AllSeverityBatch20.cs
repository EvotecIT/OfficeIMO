using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAllSeverityBatch20SecurityTests {
    [Fact]
    public void UnbackedReservedMonospaceSlotCannotSuppressAutomaticFallback() {
        string[] candidates = PdfOptions.DefaultDocumentMonospaceFontFamilyFallback
            .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(candidate => candidate.Trim())
            .ToArray();
        if (!candidates.Any(candidate => PdfEmbeddedFontFamily.TryFromSystem(candidate, out _))) {
            return;
        }

        var options = new PdfOptions();
        options.UseTextFallbacks(
            PdfTextFallbackFeatures.Default,
            new[] { PdfStandardFont.Courier },
            allowSystemFontEmbedding: true);

        Assert.True(options.HasEmbeddedStandardFontFamily(PdfStandardFont.Courier));
    }

    [Fact]
    public void ExplicitConfiguredMonospaceSlotRemainsReserved() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Courier,
            HeaderFont = PdfStandardFont.Courier,
            FooterFont = PdfStandardFont.Courier
        };

        options.UseTextFallbacks(
            PdfTextFallbackFeatures.MonospaceFont,
            new[] { PdfStandardFont.Courier },
            allowSystemFontEmbedding: true,
            preserveConfiguredFontSlots: true);

        Assert.False(options.HasEmbeddedStandardFontFamily(PdfStandardFont.Courier));
    }

    [Fact]
    public void ReservedDocumentSlotIsNotBackedWhenDocumentFallbackIsDisabled() {
        var options = new PdfOptions();

        options.UseTextFallbacks(
            PdfTextFallbackFeatures.SymbolAndEmojiFonts,
            new[] { PdfStandardFont.Helvetica },
            allowSystemFontEmbedding: true);

        Assert.False(options.HasEmbeddedStandardFontFamily(PdfStandardFont.Helvetica));
    }

    [Fact]
    public void ExplicitConfiguredDocumentSlotRemainsReserved() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            HeaderFont = PdfStandardFont.Helvetica,
            FooterFont = PdfStandardFont.Helvetica
        };

        options.UseTextFallbacks(
            PdfTextFallbackFeatures.DocumentFont,
            new[] { PdfStandardFont.Helvetica },
            allowSystemFontEmbedding: true,
            preserveConfiguredFontSlots: true);

        Assert.False(options.HasEmbeddedStandardFontFamily(PdfStandardFont.Helvetica));
    }

    [Fact]
    public void LinkedVerticalLineProducesPositiveAnnotationBounds() {
        OfficeShape line = OfficeShape.Line(0D, 0D, 0D, 20D);
        line.StrokeColor = OfficeColor.Blue;
        line.StrokeWidth = 2D;

        byte[] bytes = PdfDocument.Create()
            .Shape(line, linkUri: "https://example.test/line")
            .ToBytes();

        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).LinkAnnotations);
        Assert.True(link.X2 > link.X1);
        Assert.True(link.Y2 > link.Y1);
        Assert.True(link.X2 - link.X1 >= 2D);
    }
}
