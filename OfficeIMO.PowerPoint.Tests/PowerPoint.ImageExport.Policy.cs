using System.IO;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_StrictOmissionPolicyRejectsSkippedGroupContent() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectanglePoints(10, 10, 20, 20);
            PowerPointAutoShape second = slide.AddRectanglePoints(40, 10, 20, 20);
            slide.GroupShapes(new PowerPointShape[] { first, second }, "Omitted group");

            OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
                slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions {
                    MaxGroupShapeDepth = 0,
                    Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                }));
            Assert.Contains(exception.Diagnostics, diagnostic =>
                diagnostic.Code == "group-depth-limit" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void PowerPointSlide_StrictOmissionPolicyAllowsSmallCapsApproximation() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(160),OfficeIMO.OfficeLength.Points(90));
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox textBox = slide.AddTextBox("Keep this text",global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(120),global::OfficeIMO.OfficeLength.Points(50));
            textBox.Paragraphs[0].Runs[0].Capitalization = PowerPointCapitalization.SmallCaps;

            OfficeImageExportResult result = slide.ExportImage(
                OfficeImageExportFormat.Svg,
                new PowerPointImageExportOptions {
                    IncludeSlideBackground = false,
                    Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                });

            Assert.Contains(
                result.Diagnostics,
                diagnostic =>
                    diagnostic.Code == PowerPointImageExportDiagnosticCodes.SmallCapsApproximated &&
                    diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        }

        [Fact]
        public void PowerPointSlide_StrictOmissionPolicyRejectsUnreadableBackgroundImage() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(160),OfficeIMO.OfficeLength.Points(90));
            PowerPointSlide slide = presentation.AddSlide();
            using var emptyImage = new MemoryStream();
            slide.SetBackgroundImage(
                emptyImage,
                OfficeImageFormat.Png);

            OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
                slide.ExportImage(
                    OfficeImageExportFormat.Svg,
                    new PowerPointImageExportOptions {
                        Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                    }));

            Assert.Contains(
                exception.Diagnostics,
                diagnostic =>
                    diagnostic.Code == PowerPointImageExportDiagnosticCodes.InvalidSlideBackgroundImage &&
                    diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void PowerPointSlide_StrictOmissionPolicyRejectsUnsupportedBackgroundFill() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(160),OfficeIMO.OfficeLength.Points(90));
            PowerPointSlide slide = presentation.AddSlide();
            slide.SlidePart.Slide.CommonSlideData!.Background = new Background(
                new BackgroundProperties(
                    new A.GradientFill(
                        new A.PathGradientFill { Path = A.PathShadeValues.Shape })));

            OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
                slide.ExportImage(
                    OfficeImageExportFormat.Svg,
                    new PowerPointImageExportOptions {
                        Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                    }));

            Assert.Contains(
                exception.Diagnostics,
                diagnostic =>
                    diagnostic.Code == "unsupported-slide-background" &&
                    diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }
    }
}
