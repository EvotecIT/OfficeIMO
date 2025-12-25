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
        public void RemovingSlidesDownToOneKeepsPresentationValid() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide().AddTextBox("Slide 1");
                presentation.AddSlide().AddTextBox("Slide 2");
                presentation.AddSlide().AddTextBox("Slide 3");

                presentation.RemoveSlide(2);
                presentation.RemoveSlide(1);

                Assert.Single(presentation.Slides);
                Assert.True(presentation.DocumentIsValid);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Single(presentation.Slides);
                Assert.True(presentation.DocumentIsValid);
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
        public void RemovingLastSlideThrows() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(Path.GetTempFileName());
            presentation.AddSlide();

            Assert.Throws<InvalidOperationException>(() => presentation.RemoveSlide(0));
        }

        [Fact]
        public void MovingInvalidSlideThrows() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(Path.GetTempFileName());
            presentation.AddSlide();
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.MoveSlide(0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.MoveSlide(1, 0));
        }

        [Fact]
        public void CanDuplicateSlide() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide source = presentation.AddSlide();
                source.AddTextBox("Source slide");
                source.AddPicture(imagePath);
                source.Notes.Text = "Speaker notes";
                source.Hidden = true;

                PowerPointSlide duplicate = presentation.DuplicateSlide(0);

                Assert.Equal(2, presentation.Slides.Count);
                Assert.Equal("Source slide", duplicate.TextBoxes.First().Text);
                Assert.Single(duplicate.Pictures);
                Assert.True(duplicate.Hidden);
                Assert.Equal("Speaker notes", duplicate.Notes.Text);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Equal(2, presentation.Slides.Count);
                Assert.True(presentation.Slides[1].Hidden);
                Assert.Equal("Source slide", presentation.Slides[1].TextBoxes.First().Text);
                Assert.Equal("Speaker notes", presentation.Slides[1].Notes.Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanDuplicateSlideWithChart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide source = presentation.AddSlide();
                source.AddChart();

                presentation.DuplicateSlide(0);
                Assert.Equal(2, presentation.Slides.Count);
                Assert.Single(presentation.Slides[0].Charts);
                Assert.Single(presentation.Slides[1].Charts);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Equal(2, presentation.Slides.Count);
                Assert.Single(presentation.Slides[0].Charts);
                Assert.Single(presentation.Slides[1].Charts);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanImportSlideFromAnotherPresentation() {
            string sourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string targetPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation source = PowerPointPresentation.Create(sourcePath)) {
                PowerPointSlide sourceSlide = source.AddSlide();
                sourceSlide.AddTextBox("Imported slide");
                sourceSlide.AddPicture(imagePath);
                sourceSlide.Notes.Text = "Imported notes";
                sourceSlide.Hidden = true;
                source.Save();

                using (PowerPointPresentation target = PowerPointPresentation.Create(targetPath)) {
                    PowerPointSlide imported = target.ImportSlide(source, 0);

                    Assert.Single(target.Slides);
                    Assert.Equal("Imported slide", imported.TextBoxes.First().Text);
                    Assert.Single(imported.Pictures);
                    Assert.True(imported.Hidden);
                    Assert.Equal("Imported notes", imported.Notes.Text);

                    target.Save();
                }
            }

            using (PowerPointPresentation target = PowerPointPresentation.Open(targetPath)) {
                Assert.Single(target.Slides);
                Assert.Equal("Imported slide", target.Slides[0].TextBoxes.First().Text);
                Assert.Single(target.Slides[0].Pictures);
                Assert.True(target.Slides[0].Hidden);
                Assert.Equal("Imported notes", target.Slides[0].Notes.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }
    }
}
