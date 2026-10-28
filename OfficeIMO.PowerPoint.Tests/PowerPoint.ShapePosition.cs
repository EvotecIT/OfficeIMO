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

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox textBox = slide.AddTextBox("Hello", left, top, width, height);
                PowerPointPicture picture = slide.AddImage(imagePath, left, top, width, height);
                PowerPointTable table = slide.AddTable(2,2,global::OfficeIMO.OfficeLength.Emus(left),global::OfficeIMO.OfficeLength.Emus(top),global::OfficeIMO.OfficeLength.Emus(width),global::OfficeIMO.OfficeLength.Emus(height));

                Assert.Equal(left, textBox.Left.ToEmus());
                Assert.Equal(top, textBox.Top.ToEmus());
                Assert.Equal(width, textBox.Width.ToEmus());
                Assert.Equal(height, textBox.Height.ToEmus());

                textBox.Left+= OfficeIMO.OfficeLength.Emus(1000);
                picture.Top+= OfficeIMO.OfficeLength.Emus(2000);
                table.Width+= OfficeIMO.OfficeLength.Emus(3000);

                presentation.Save();
            }

            using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                PowerPointTextBox textBox = slide.TextBoxes.First();
                PowerPointPicture picture = slide.Pictures.First();
                PowerPointTable table = slide.Tables.First();

                Assert.Equal(left + 1000, textBox.Left.ToEmus());
                Assert.Equal(top, textBox.Top.ToEmus());
                Assert.Equal(top + 2000, picture.Top.ToEmus());
                Assert.Equal(left, picture.Left.ToEmus());
                Assert.Equal(width + 3000, table.Width.ToEmus());
                Assert.Equal(height, table.Height.ToEmus());
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanUseCentimeterHelpers() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox textBox = slide.AddTextBox("Hello");

                textBox.SetPosition(OfficeIMO.OfficeLength.Centimeters(2.5),OfficeIMO.OfficeLength.Centimeters(3.5));
                textBox.SetSize(OfficeIMO.OfficeLength.Centimeters(10),OfficeIMO.OfficeLength.Centimeters(4));

                Assert.Equal(PowerPointUnits.FromCentimeters(2.5), textBox.Left.ToEmus());
                Assert.Equal(PowerPointUnits.FromCentimeters(3.5), textBox.Top.ToEmus());
                Assert.Equal(PowerPointUnits.FromCentimeters(10), textBox.Width.ToEmus());
                Assert.Equal(PowerPointUnits.FromCentimeters(4), textBox.Height.ToEmus());

                presentation.Save();
            }

            using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                PowerPointTextBox textBox = slide.TextBoxes.First();

                Assert.InRange(textBox.Left.ToCentimeters(), 2.49, 2.51);
                Assert.InRange(textBox.Top.ToCentimeters(), 3.49, 3.51);
                Assert.InRange(textBox.Width.ToCentimeters(), 9.99, 10.01);
                Assert.InRange(textBox.Height.ToCentimeters(), 3.99, 4.01);
            }

            File.Delete(filePath);
        }
    }
}
