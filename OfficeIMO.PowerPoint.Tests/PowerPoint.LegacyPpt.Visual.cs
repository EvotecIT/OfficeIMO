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
            using PowerPointDocument presentation = PowerPointDocument.Load(
                ShapeFixturePath);
            var options = new PowerPointImageExportOptions {
                BackgroundColor = OfficeColor.White,
                IncludeSlideBackground = false
            };

            PowerPointSlideVisualInfo snapshot = presentation.Slides[0]
                .CreateVisualSnapshot(options);
            OfficeImageExportDiagnostic[] unexpectedDiagnostics = snapshot.Diagnostics.Where(diagnostic =>
                (diagnostic.Severity != OfficeImageExportDiagnosticSeverity.Info
                    || diagnostic.LossKind != OfficeConversionLossKind.None)
                && diagnostic.Code != OfficeImageExportDiagnosticCodes.FontSubstituted
                && diagnostic.Code != "font-family-substitution").ToArray();

            byte[] actual = presentation.Slides[0].ToImage(options);
            VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                File.ReadAllBytes(ShapeVisualBaselinePath), actual,
                channelTolerance: 16,
                allowedDifferentPixels: 20000,
                maximumMeanAbsoluteError: 8D,
                maximumRootMeanSquareError: 30D,
                maximumMeanLuminanceError: 10D);

            Assert.True(comparison.Passed && unexpectedDiagnostics.Length == 0,
                $"Binary shape raster comparison passed={comparison.Passed}: " +
                $"{comparison.DifferentPixels} of {comparison.TotalPixels} pixels " +
                $"(maximum channel delta {comparison.MaxChannelDelta}, " +
                $"MAE {comparison.MeanAbsoluteError:F3}, RMSE {comparison.RootMeanSquareError:F3}, " +
                $"luminance MAE {comparison.MeanLuminanceError:F3}). " +
                "Unexpected snapshot diagnostics: " + string.Join("; ", unexpectedDiagnostics.Select(diagnostic =>
                    $"{diagnostic.Code} [{diagnostic.LossKind}]: {diagnostic.Message}")));
        }

        [Fact]
        public void MicrosoftAuthoredImport_DoesNotRenderMasterPlaceholderPrompts() {
            using PowerPointDocument presentation = PowerPointDocument.Load(
                AccessibilityFixturePath);
            var options = new PowerPointImageExportOptions {
                BackgroundColor = OfficeColor.White,
                IncludeSlideBackground = false
            };

            PowerPointSlideVisualInfo snapshot = presentation.Slides[0]
                .CreateVisualSnapshot(options);
            Assert.NotEmpty(snapshot.TextSpans);
            Assert.DoesNotContain("Click to edit", string.Concat(snapshot.TextSpans.Select(span => span.Text)),
                StringComparison.Ordinal);
            OfficeImageExportDiagnostic[] unexpectedDiagnostics = snapshot.Diagnostics.Where(diagnostic =>
                (diagnostic.Severity != OfficeImageExportDiagnosticSeverity.Info
                    || diagnostic.LossKind != OfficeConversionLossKind.None)
                && diagnostic.Code != OfficeImageExportDiagnosticCodes.FontSubstituted
                && diagnostic.Code != "font-family-substitution").ToArray();

            byte[] actual = presentation.Slides[0].ToImage(options);
            VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                File.ReadAllBytes(AccessibilityVisualBaselinePath), actual,
                channelTolerance: 16,
                allowedDifferentPixels: 8000,
                maximumMeanAbsoluteError: 4D,
                maximumRootMeanSquareError: 22D,
                maximumMeanLuminanceError: 6D);

            Assert.True(comparison.Passed && unexpectedDiagnostics.Length == 0,
                $"Microsoft-authored raster comparison passed={comparison.Passed}: " +
                $"{comparison.DifferentPixels} of {comparison.TotalPixels} pixels " +
                $"(maximum channel delta {comparison.MaxChannelDelta}, " +
                $"MAE {comparison.MeanAbsoluteError:F3}, RMSE {comparison.RootMeanSquareError:F3}, " +
                $"luminance MAE {comparison.MeanLuminanceError:F3}). " +
                "Unexpected snapshot diagnostics: " + string.Join("; ", unexpectedDiagnostics.Select(diagnostic =>
                    $"{diagnostic.Code} [{diagnostic.LossKind}]: {diagnostic.Message}")));
        }
    }
}
