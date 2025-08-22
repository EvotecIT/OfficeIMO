using System.IO;
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
            paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);

            var loc = paragraph.Image!.Location;
            Assert.Equal(0, loc.X);
            Assert.Equal(0, loc.Y);

            document.Save(false);
        }

        [Fact]
        public void Test_ImageLocation_Floating() {
            var filePath = Path.Combine(_directoryWithFiles, "ImageLocationFloating.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph();
            paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50, WrapTextImage.Square);
            int offset = 914400;
            var image = paragraph.Image!;
            image.horizontalPosition = new HorizontalPosition() {
                RelativeFrom = HorizontalRelativePositionValues.Page,
                PositionOffset = new PositionOffset { Text = offset.ToString() }
            };
            image.verticalPosition = new VerticalPosition() {
                RelativeFrom = VerticalRelativePositionValues.Page,
                PositionOffset = new PositionOffset { Text = offset.ToString() }
            };

            var loc = image.Location;
            Assert.Equal(offset, loc.X);
            Assert.Equal(offset, loc.Y);

            document.Save(false);
        }
    }
}
