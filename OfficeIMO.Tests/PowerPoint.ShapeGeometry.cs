using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeGeometry {
        [Fact]
        public void CanSetShapePositionAndSize() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            long left = 1000000L;
            long top = 2000000L;
            long width = 3000000L;
            long height = 4000000L;

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBox("Test", left, top, width, height);
                slide.AddPicture(imagePath, left, top, width, height);
                slide.AddTable(2, 2, left, top, width, height);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointSlide slide = presentation.Slides[0];
                PPTextBox box = slide.TextBoxes.First();
                Assert.Equal(left, box.Left);
                Assert.Equal(top, box.Top);
                Assert.Equal(width, box.Width);
                Assert.Equal(height, box.Height);

                PPPicture pic = slide.Pictures.First();
                Assert.Equal(left, pic.Left);
                Assert.Equal(top, pic.Top);
                Assert.Equal(width, pic.Width);
                Assert.Equal(height, pic.Height);

                PPTable tbl = slide.Tables.First();
                Assert.Equal(left, tbl.Left);
                Assert.Equal(top, tbl.Top);
                Assert.Equal(width, tbl.Width);
                Assert.Equal(height, tbl.Height);
            }

            File.Delete(filePath);
        }
    }
}
