using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointFluentPresentation {
        [Fact(Skip = "Doesn't work after changes to PowerPoint")]
        public void CanBuildPresentationFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AsFluent()
                    .Slide(0, 1)
                        .Title("Fluent Title")
                        .TextBox("Hello")
                        .Bullets("One", "Two")
                        .Image(imagePath)
                        .Table(2, 2)
                        .Notes("Notes text")
                        .End()
                    .Slide(s => s.Title("Second Slide"))
                    .End()
                    .Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Equal(2, presentation.Slides.Count);
                PowerPointSlide slide = presentation.Slides[0];
                Assert.Equal("Notes text", slide.Notes.Text);
                Assert.Equal(5, slide.Shapes.Count);
                var textBoxes = slide.Shapes.OfType<PowerPointTextBox>().ToList();
                Assert.Equal(3, textBoxes.Count);
                Assert.Equal("Fluent Title", textBoxes[0].Text);
                Assert.Equal("Hello", textBoxes[1].Text);
                Assert.Equal(1, slide.LayoutIndex);

                PowerPointSlide slide2 = presentation.Slides[1];
                var textBoxes2 = slide2.Shapes.OfType<PowerPointTextBox>().ToList();
                Assert.Single(textBoxes2);
                Assert.Equal("Second Slide", textBoxes2[0].Text);
            }

            File.Delete(filePath);
        }
    }
}