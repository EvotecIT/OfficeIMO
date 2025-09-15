using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ImageLocation_Inline() {
            var filePath = Path.Combine(_directoryWithFiles, "ImageLocationInline.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph();
            var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Kulek.jpg");
            paragraph.AddImage(imagePath, 50, 50);

            var image = document.Images.FirstOrDefault();
            Assert.NotNull(image);
            var loc = image!.Location;
            Assert.Equal(0, loc.X);
            Assert.Equal(0, loc.Y);

            document.Save(false);
        }

        [Fact]
        public void Test_ImageLocation_Floating() {
            var filePath = Path.Combine(_directoryWithFiles, "ImageLocationFloating.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph();
            var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Kulek.jpg");
            paragraph.AddImage(imagePath, 50, 50, WrapTextImage.Square);
            var paraImage = paragraph.Image;
            Assert.NotNull(paraImage);
            int offset = 914400;
            paraImage!.horizontalPosition = new HorizontalPosition() {
                RelativeFrom = HorizontalRelativePositionValues.Page,
                PositionOffset = new PositionOffset { Text = offset.ToString() }
            };
            paraImage.verticalPosition = new VerticalPosition() {
                RelativeFrom = VerticalRelativePositionValues.Page,
                PositionOffset = new PositionOffset { Text = offset.ToString() }
            };

            var image = document.Images.FirstOrDefault();
            Assert.NotNull(image);
            var loc = image!.Location;
            Assert.Equal(offset, loc.X);
            Assert.Equal(offset, loc.Y);

            document.Save(false);
        }
    }
}
