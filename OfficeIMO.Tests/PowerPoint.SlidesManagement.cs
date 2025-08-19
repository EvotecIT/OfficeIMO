using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointSlidesManagement {
        [Fact]
        public void CanRemoveSlide() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide().AddTextBox("Slide 1");
                presentation.AddSlide().AddTextBox("Slide 2");
                presentation.AddSlide().AddTextBox("Slide 3");

                presentation.RemoveSlide(1);

                Assert.Equal(2, presentation.Slides.Count);
                Assert.Equal("Slide 1", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Slide 3", presentation.Slides[1].TextBoxes.First().Text);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Equal(2, presentation.Slides.Count);
                Assert.Equal("Slide 1", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Slide 3", presentation.Slides[1].TextBoxes.First().Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanMoveSlide() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide().AddTextBox("Slide 1");
                presentation.AddSlide().AddTextBox("Slide 2");
                presentation.AddSlide().AddTextBox("Slide 3");

                presentation.MoveSlide(0, 2);

                Assert.Equal("Slide 2", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Slide 3", presentation.Slides[1].TextBoxes.First().Text);
                Assert.Equal("Slide 1", presentation.Slides[2].TextBoxes.First().Text);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Equal(3, presentation.Slides.Count);
                Assert.Equal("Slide 2", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Slide 3", presentation.Slides[1].TextBoxes.First().Text);
                Assert.Equal("Slide 1", presentation.Slides[2].TextBoxes.First().Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RemovingInvalidSlideThrows() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(Path.GetTempFileName());
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.RemoveSlide(0));
        }

        [Fact]
        public void MovingInvalidSlideThrows() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(Path.GetTempFileName());
            presentation.AddSlide();
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.MoveSlide(0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.MoveSlide(1, 0));
        }
    }
}
