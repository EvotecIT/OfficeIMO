using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointTextAutoFitTests {
        [Fact]
        public void CanSetTextAutoFitOptions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Auto-fit me");
                box.SetTextAutoFit(PowerPointTextAutoFit.Normal,
                    new PowerPointTextAutoFitOptions(fontScalePercent: 85, lineSpaceReductionPercent: 10));
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                Shape shape = slidePart.Slide.Descendants<Shape>().First();
                A.BodyProperties body = shape.TextBody!.GetFirstChild<A.BodyProperties>()!;
                A.NormalAutoFit normal = body.GetFirstChild<A.NormalAutoFit>()!;
                Assert.Equal(85000, normal.FontScale!.Value);
                Assert.Equal(10000, normal.LineSpaceReduction!.Value);
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointTextBox box = presentation.Slides.First().TextBoxes.First();
                Assert.Equal(PowerPointTextAutoFit.Normal, box.TextAutoFit);
                PowerPointTextAutoFitOptions? options = box.TextAutoFitOptions;
                Assert.NotNull(options);
                Assert.Equal(85d, options!.Value.FontScalePercent!.Value, 3);
                Assert.Equal(10d, options.Value.LineSpaceReductionPercent!.Value, 3);
            }

            File.Delete(filePath);
        }
    }
}
