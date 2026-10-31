using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfAuthoringLayoutDiagnosticsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AlternatingGlyphContoursMatchIndependentPaintDiagnostics(bool transformed) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf"));
        var items = new[] {
            (Text: "B8ig. AQ", X: 28D, Y: 28D, Size: 9D),
            (Text: "Q.Ag 8Bi", X: 120D, Y: 48D, Size: 18D),
            (Text: "B8ig. AQ", X: 28D, Y: 78D, Size: 9D)
        };
        (byte[] Bytes, string[] Warnings) Render(int first, int count, bool diagnose) {
            var report = new PdfConversionReport();
            var options = new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(260), PageHeight = global::OfficeIMO.OfficeLength.Points(180), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9),
                CompressEmbeddedFonts = false, CompressContentStreams = false
            }.EmbedStandardFont(PdfStandardFont.Helvetica, font, "Alternating contour probe");
            if (diagnose) options.ReportDiagnosticsTo(report, "AlternatingGlyphProbe");
            byte[] bytes = PdfDocument.Create(options).Canvas(canvas => canvas.Effect(
                transformed ? OfficeTransform.RotateDegrees(15, 24, 24) : OfficeTransform.Identity, 1D,
                effect => effect.Clip(global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(100), child => {
                    for (int i = first; i < first + count; i++) {
                        var item = items[i];
                        child.Text(item.Text, global::OfficeIMO.OfficeLength.Points(item.X), global::OfficeIMO.OfficeLength.Points(item.Y), global::OfficeIMO.OfficeLength.Points(90), global::OfficeIMO.OfficeLength.Points(28), fontSize: global::OfficeIMO.OfficeLength.Points(item.Size));
                    }
                }))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
            return (bytes, report.Warnings.Select(warning => System.Text.Json.JsonSerializer.Serialize(warning)).ToArray());
        }

        var combined = Render(0, items.Length, true);
        Assert.Equal(Render(0, items.Length, false).Bytes, combined.Bytes);
        // Separate documents each emit their own document-wide notices. A single
        // document deduplicates those notices; keep every distinct occurrence,
        // including its exact source bounds, rather than multiplying setup notices.
        string[] independent = Enumerable.Range(0, items.Length).SelectMany(i => Render(i, 1, true).Warnings)
            .Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToArray();
        Assert.NotEmpty(independent);
        Assert.Equal(independent, combined.Warnings.OrderBy(value => value, StringComparer.Ordinal).ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedEmbeddedGlyphsRetainIndependentPlacementAndClipDiagnostics(bool transformed) {
        byte[] regular = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf"));
        byte[] bold = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Bold.ttf"));
        var report = new PdfConversionReport();
        byte[] Render(bool diagnose) {
            var options = new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(260), PageHeight = global::OfficeIMO.OfficeLength.Points(180), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9),
                CompressEmbeddedFonts = false, CompressContentStreams = false
            };
            options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Carlito", regular, bold));
            if (diagnose) options.ReportDiagnosticsTo(report, "RepeatedGlyphProbe");
            return PdfDocument.Create(options).Canvas(canvas => canvas.Effect(
                transformed ? OfficeTransform.RotateDegrees(15, 24, 24) : OfficeTransform.Identity, 1D,
                effect => effect.Clip(global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(60), child => {
                    child.Text("AAAA", global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(18));
                    child.Text(new[] { PdfTextRun.Bolded("AAAA", fontSize: global::OfficeIMO.OfficeLength.Points(9)) }, global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(18));
                    child.Text(new[] { PdfTextRun.Normal("AAAA", fontSize: global::OfficeIMO.OfficeLength.Points(18)) }, global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(52), global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(28));
                    child.Text("AAAA", global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(18));
                }))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        Assert.Equal(Render(false), Render(true));
        Assert.Single(report.Warnings, warning => warning.Code == "canvas-text-canvas-clip-paint-clipped");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code.EndsWith("PaintBoundsUnmeasured", StringComparison.Ordinal));
        Assert.True(report.HasLoss);
    }

    [Fact]
    public void ContainedRepeatedEmbeddedGlyphsPreserveBytesAndHaveNoPaintLoss() {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf"));
        var report = new PdfConversionReport();
        byte[] Render(bool diagnose) {
            var options = new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(260), PageHeight = global::OfficeIMO.OfficeLength.Points(180), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9),
                CompressEmbeddedFonts = false, CompressContentStreams = false
            };
            if (diagnose) options.ReportDiagnosticsTo(report, "ContainedGlyphProbe");
            options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "Contained glyph probe");
            return PdfDocument.Create(options).Canvas(canvas => {
                canvas.Text("AAAA AAAA", global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(18));
                canvas.Text("AAAA AAAA", global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(58), global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(18));
            }).ToBytes();
        }

        Assert.Equal(Render(false), Render(true));
        Assert.DoesNotContain(report.Warnings, warning => warning.Code.Contains("Paint", StringComparison.Ordinal));
        Assert.False(report.HasLoss);
    }
}
