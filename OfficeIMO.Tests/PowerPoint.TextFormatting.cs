using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointTextFormatting {
        [Fact]
        public void CanApplyFormattingToTextBoxAndBullets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Hello");
                box.Bold = true;
                box.Italic = true;
                box.FontSize = 24;
                box.FontName = "Arial";
                box.Color = "FF0000";
                box.AddBullet("Bullet1");
                box.AddBullet("Bullet2");
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                Shape shape = slidePart.Slide.Descendants<Shape>().First();
                var paragraphs = shape.TextBody!.Elements<A.Paragraph>().ToList();
                foreach (var paragraph in paragraphs) {
                    A.Run run = paragraph.GetFirstChild<A.Run>()!;
                    A.RunProperties rp = run.RunProperties!;
                    Assert.True(rp.Bold == true);
                    Assert.True(rp.Italic == true);
                    Assert.Equal(2400, rp.FontSize!.Value);
                    Assert.Equal("Arial", rp.GetFirstChild<A.LatinFont>()?.Typeface);
                    Assert.Equal("FF0000", rp.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val);
                }
            }

            File.Delete(filePath);
        }
    }
}
