using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
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
            }

            File.Delete(filePath);
        }
    }
}
