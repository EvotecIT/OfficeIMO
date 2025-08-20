using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointBasicDocument {
        [Fact]
        public void CanCreateSaveAndLoadPresentation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox text = slide.AddTextBox("Hello");
                text.AddBullet("Bullet1");
                slide.AddPicture(imagePath);
                slide.AddTable(2, 2);
                slide.Notes.Text = "Test notes";
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Single(presentation.Slides);
                PowerPointSlide slide = presentation.Slides[0];
                PowerPointTextBox box = slide.Shapes.OfType<PowerPointTextBox>().First();
                Assert.Equal("Hello", box.Text);
                Assert.Equal("Test notes", slide.Notes.Text);
                Assert.Equal(3, slide.Shapes.Count); // textbox, picture, table
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                Assert.NotNull(document.CoreFilePropertiesPart);
                Assert.NotNull(document.ExtendedFilePropertiesPart);
                PresentationPart part = document.PresentationPart!;
                Assert.NotNull(part.PresentationPropertiesPart);
                Assert.NotNull(part.ViewPropertiesPart);
                Assert.NotNull(part.TableStylesPart);

                SlidePart slidePart = part.SlideParts.First();
                ShapeTree tree = slidePart.Slide.CommonSlideData!.ShapeTree!;
                Assert.NotNull(tree.GetFirstChild<NonVisualGroupShapeProperties>());
                Assert.NotNull(tree.GetFirstChild<GroupShapeProperties>());

                var ids = tree.Descendants<NonVisualDrawingProperties>().Select(dp => dp.Id!.Value).ToList();
                Assert.Equal(ids.Count, ids.Distinct().Count());
                Assert.Contains(1U, ids);
                Assert.Contains(2U, ids);
                Assert.Contains(3U, ids);
                Assert.Contains(4U, ids);
            }

            File.Delete(filePath);
        }
    }
}
