using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapePosition {
        [Fact]
        public void CanSetShapePositions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            long left = 1000000L;
            long top = 2000000L;
            long width = 3000000L;
            long height = 4000000L;

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox textBox = slide.AddTextBox("Hello", left, top, width, height);
                PowerPointPicture picture = slide.AddPicture(imagePath, left, top, width, height);
                PowerPointTable table = slide.AddTable(2, 2, left, top, width, height);

                Assert.Equal(left, textBox.Left);
                Assert.Equal(top, textBox.Top);
                Assert.Equal(width, textBox.Width);
                Assert.Equal(height, textBox.Height);

                textBox.Left += 1000;
                picture.Top += 2000;
                table.Width += 3000;

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                PowerPointTextBox textBox = slide.TextBoxes.First();
                PowerPointPicture picture = slide.Pictures.First();
                PowerPointTable table = slide.Tables.First();

                Assert.Equal(left + 1000, textBox.Left);
                Assert.Equal(top, textBox.Top);
                Assert.Equal(top + 2000, picture.Top);
                Assert.Equal(left, picture.Left);
                Assert.Equal(width + 3000, table.Width);
                Assert.Equal(height, table.Height);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanUseCentimeterHelpers() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox textBox = slide.AddTextBox("Hello");

                textBox.SetPositionCm(2.5, 3.5);
                textBox.SetSizeCm(10, 4);

                Assert.Equal(PowerPointUnits.Cm(2.5), textBox.Left);
                Assert.Equal(PowerPointUnits.Cm(3.5), textBox.Top);
                Assert.Equal(PowerPointUnits.Cm(10), textBox.Width);
                Assert.Equal(PowerPointUnits.Cm(4), textBox.Height);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                PowerPointTextBox textBox = slide.TextBoxes.First();

                Assert.InRange(textBox.LeftCm, 2.49, 2.51);
                Assert.InRange(textBox.TopCm, 3.49, 3.51);
                Assert.InRange(textBox.WidthCm, 9.99, 10.01);
                Assert.InRange(textBox.HeightCm, 3.99, 4.01);
            }

            File.Delete(filePath);
        }
    }
}
