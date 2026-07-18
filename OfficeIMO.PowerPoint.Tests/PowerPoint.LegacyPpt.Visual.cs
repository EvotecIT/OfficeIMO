using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        private static string ShapeVisualBaselinePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "VisualBaselines",
            "ShapePowerPoint.libreoffice.png");

        private static string AccessibilityVisualBaselinePath => Path.Combine(
            AppContext.BaseDirectory, "Documents", "LegacyPptCorpus", "VisualBaselines",
            "AccessibilityPowerPoint.libreoffice.png");

        [Fact]
        public void ShapeImport_MatchesLibreOfficeVisualReferenceWithinDocumentedTolerance() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                ShapeFixturePath);
            var options = new PowerPointImageExportOptions {
                BackgroundColor = OfficeColor.White,
                IncludeSlideBackground = false
            };

            PowerPointSlideVisualSnapshot snapshot = presentation.Slides[0]
                .CreateVisualSnapshot(options);
            Assert.True(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>()
                .Count(shape => shape.Shape.FillGradient != null) >= 7);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic =>
                diagnostic.Code != OfficeImageExportDiagnosticCodes.FontSubstituted);

            byte[] actual = presentation.Slides[0].ToPng(options);
            VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                File.ReadAllBytes(ShapeVisualBaselinePath), actual,
                channelTolerance: 16,
                allowedDifferentPixels: 20000,
                maximumMeanAbsoluteError: 8D,
                maximumRootMeanSquareError: 30D,
                maximumMeanLuminanceError: 10D);

            Assert.True(comparison.Passed,
                $"Binary shape rendering differs from the LibreOffice reference at " +
                $"{comparison.DifferentPixels} of {comparison.TotalPixels} pixels " +
                $"(maximum channel delta {comparison.MaxChannelDelta}, " +
                $"MAE {comparison.MeanAbsoluteError:F3}, RMSE {comparison.RootMeanSquareError:F3}, " +
                $"luminance MAE {comparison.MeanLuminanceError:F3}).");
        }

        [Fact]
        public void MicrosoftAuthoredImport_DoesNotRenderMasterPlaceholderPrompts() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AccessibilityFixturePath);
            var options = new PowerPointImageExportOptions {
                BackgroundColor = OfficeColor.White,
                IncludeSlideBackground = false
            };

            PowerPointSlideVisualSnapshot snapshot = presentation.Slides[0]
                .CreateVisualSnapshot(options);
            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text =>
                text.Text.Contains("Click to edit", StringComparison.Ordinal));
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic =>
                diagnostic.Code != OfficeImageExportDiagnosticCodes.FontSubstituted);

            byte[] actual = presentation.Slides[0].ToPng(options);
            VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                File.ReadAllBytes(AccessibilityVisualBaselinePath), actual,
                channelTolerance: 16,
                allowedDifferentPixels: 8000,
                maximumMeanAbsoluteError: 4D,
                maximumRootMeanSquareError: 22D,
                maximumMeanLuminanceError: 6D);

            Assert.True(comparison.Passed,
                $"Microsoft-authored binary rendering differs from the LibreOffice reference at " +
                $"{comparison.DifferentPixels} of {comparison.TotalPixels} pixels " +
                $"(maximum channel delta {comparison.MaxChannelDelta}, " +
                $"MAE {comparison.MeanAbsoluteError:F3}, RMSE {comparison.RootMeanSquareError:F3}, " +
                $"luminance MAE {comparison.MeanLuminanceError:F3}).");
        }
    }
}
