using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointFluentPresentation {
        [Fact]
        public void CanBuildPresentationFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AsFluent()
                    .Slide()
                        .Title("Fluent Title")
                        .Text("Hello")
                        .Bullets("One", "Two")
                        .Image(imagePath)
                        .Table(2, 2)
                        .Notes("Notes text");
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Single(presentation.Slides);
                PowerPointSlide slide = presentation.Slides[0];
                Assert.Equal("Notes text", slide.Notes.Text);
                Assert.Equal(5, slide.Shapes.Count);
                var textBoxes = slide.Shapes.OfType<PPTextBox>().ToList();
                Assert.Equal(3, textBoxes.Count);
                Assert.Equal("Fluent Title", textBoxes[0].Text);
                Assert.Equal("Hello", textBoxes[1].Text);
            }

            File.Delete(filePath);
        }
    }
}