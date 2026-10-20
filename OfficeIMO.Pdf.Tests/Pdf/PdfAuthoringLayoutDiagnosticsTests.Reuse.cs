using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfAuthoringLayoutDiagnosticsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedEmbeddedGlyphsRetainIndependentPlacementAndClipDiagnostics(bool transformed) {
        byte[] regular = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf"));
        byte[] bold = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Bold.ttf"));
        var report = new PdfConversionReport();
        byte[] Render(bool diagnose) {
            var options = new PdfOptions {
                PageWidth = 260, PageHeight = 180, MarginLeft = 24, MarginRight = 24,
                MarginTop = 24, MarginBottom = 24, DefaultFontSize = 9,
                CompressEmbeddedFonts = false, CompressContentStreams = false
            };
            options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Carlito", regular, bold));
            if (diagnose) options.ReportDiagnosticsTo(report, "RepeatedGlyphProbe");
            return PdfDocument.Create(options).Canvas(canvas => canvas.Effect(
                transformed ? OfficeTransform.RotateDegrees(15, 24, 24) : OfficeTransform.Identity, 1D,
                effect => effect.Clip(24, 24, 100, 60, child => {
                    child.Text("AAAA", 28, 28, 80, 18);
                    child.Text(new[] { PdfTextRun.Bolded("AAAA", fontSize: 9) }, 28, 40, 80, 18);
                    child.Text(new[] { PdfTextRun.Normal("AAAA", fontSize: 18) }, 28, 52, 80, 28);
                    child.Text("AAAA", 120, 28, 80, 18);
                }))).ToBytes();
        }

        Assert.Equal(Render(false), Render(true));
        Assert.Single(report.Warnings, warning => warning.Code == "CanvasTextCanvasClipPaintClipped");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code.EndsWith("PaintBoundsUnmeasured", StringComparison.Ordinal));
        Assert.True(report.HasLoss);
    }

    [Fact]
    public void ContainedRepeatedEmbeddedGlyphsPreserveBytesAndHaveNoPaintLoss() {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf"));
        var report = new PdfConversionReport();
        byte[] Render(bool diagnose) {
            var options = new PdfOptions {
                PageWidth = 260, PageHeight = 180, MarginLeft = 24, MarginRight = 24,
                MarginTop = 24, MarginBottom = 24, DefaultFontSize = 9,
                CompressEmbeddedFonts = false, CompressContentStreams = false
            };
            if (diagnose) options.ReportDiagnosticsTo(report, "ContainedGlyphProbe");
            options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "Contained glyph probe");
            return PdfDocument.Create(options).Canvas(canvas => {
                canvas.Text("AAAA AAAA", 28, 28, 120, 18);
                canvas.Text("AAAA AAAA", 28, 58, 120, 18);
            }).ToBytes();
        }

        Assert.Equal(Render(false), Render(true));
        Assert.DoesNotContain(report.Warnings, warning => warning.Code.Contains("Paint", StringComparison.Ordinal));
        Assert.False(report.HasLoss);
    }
}
