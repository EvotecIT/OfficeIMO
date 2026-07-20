using System.IO;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_StrictOmissionPolicyRejectsSkippedGradient() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape source = slide.AddShapePoints(
                A.ShapeTypeValues.Rectangle,
                30,
                25,
                100,
                40);
            AddShapePathGradient(
                Assert.IsType<Shape>(source.Element),
                A.PathShadeValues.Shape);

            OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
                slide.ExportImage(
                    OfficeImageExportFormat.Svg,
                    new PowerPointImageExportOptions {
                        IncludeSlideBackground = false,
                        Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                    }));

            Assert.Contains(
                exception.Diagnostics,
                diagnostic =>
                    diagnostic.Code == PowerPointImageExportDiagnosticCodes.UnsupportedShape &&
                    diagnostic.Message.Contains("gradient", StringComparison.OrdinalIgnoreCase) &&
                    diagnostic.LossKind == OfficeImageExportLossKind.Omission);
        }

        [Fact]
        public void PowerPointSlide_StrictOmissionPolicyAllowsRectangularFrameApproximation() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 90);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox textBox = slide.AddTextBoxPoints("Keep this text", 20, 20, 120, 50);
            Shape shape = Assert.IsType<Shape>(textBox.Element);
            shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset = A.ShapeTypeValues.Funnel;

            OfficeImageExportResult result = slide.ExportImage(
                OfficeImageExportFormat.Svg,
                new PowerPointImageExportOptions {
                    IncludeSlideBackground = false,
                    Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                });

            Assert.Contains(
                result.Diagnostics,
                diagnostic =>
                    diagnostic.Code == PowerPointImageExportDiagnosticCodes.UnsupportedShape &&
                    diagnostic.Message.Contains("frame geometry", StringComparison.OrdinalIgnoreCase) &&
                    diagnostic.LossKind == OfficeImageExportLossKind.Approximation);
        }
    }
}
