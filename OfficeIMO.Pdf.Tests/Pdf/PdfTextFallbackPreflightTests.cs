using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextFallbackPreflightTests {
    [Fact]
    public void WinAnsiOnlyTextDoesNotResolveAutomaticEmbeddedFonts() {
        PdfTextFallbackFeatures resolved = PdfTextDiagnostics.ResolveRequiredFallbackFeatures(
            PdfTextFallbackFeatures.Default,
            new[] { "ASCII", "Résumé – approved", "Line one\nLine two" });

        Assert.Equal(PdfTextFallbackFeatures.None, resolved);
    }

    [Fact]
    public void NonWinAnsiTextRetainsEveryRequestedFallbackGroup() {
        PdfTextFallbackFeatures requested =
            PdfTextFallbackFeatures.DocumentFont |
            PdfTextFallbackFeatures.SymbolAndEmojiFonts |
            PdfTextFallbackFeatures.MultilingualFonts;

        PdfTextFallbackFeatures resolved = PdfTextDiagnostics.ResolveRequiredFallbackFeatures(
            requested,
            new[] { "Zażółć gęślą jaźń" });

        Assert.Equal(requested, resolved);
    }
}
