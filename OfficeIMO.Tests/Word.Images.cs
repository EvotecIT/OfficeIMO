using System.IO;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;
using Xunit;

using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_CreatingWordDocumentWithImages() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImages.docx");
            using var document = WordDocument.Create(filePath);

            document.BuiltinDocumentProperties.Title = "This is sparta";
            document.BuiltinDocumentProperties.Creator = "Przemek";

            var paragraph = document.AddParagraph("This paragraph starts with some text");
            paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

            // lets add image to paragraph
            paragraph.AddImage(Path.Combine(_directoryWithImages, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22);
            //paragraph.Image.WrapText = true; // WrapSideValues.Both;

            var paragraph5 = paragraph.AddText("and more text");
            paragraph5.Bold = true;

            Assert.Equal(22d, paragraph.Image.Height!.Value, 15);
            Assert.Equal(22d, paragraph.Image.Width!.Value, 15);

            document.AddParagraph("This adds another picture with 500x500");

            var filePathImage = Path.Combine(_directoryWithImages, "Kulek.jpg");
            var paragraph2 = document.AddParagraph();
            paragraph2.AddImage(filePathImage, 500, 500);
            //paragraph2.Image.BlackWiteMode = BlackWhiteModeValues.GrayWhite;
            paragraph2.Image.Rotation = 180;
            paragraph2.Image.Shape = ShapeTypeValues.ActionButtonMovie;

            Assert.Equal(500d, paragraph2.Image.Height!.Value, 15);
            Assert.Equal(500d, paragraph2.Image.Width!.Value, 15);

            document.AddParagraph("This adds another picture with 100x100");

            var paragraph3 = document.AddParagraph();
            paragraph3.AddImage(filePathImage, 100, 100);

            // we add paragraph with an image
            var paragraph4 = document.AddParagraph();
            paragraph4.AddImage(filePathImage);

            // we can also overwrite height later on
            paragraph4.Image.Height = 50;
            paragraph4.Image.Width = 50;
            // this doesn't work
            paragraph4.Image.HorizontalFlip = true;


            Assert.Equal(50d, paragraph4.Image.Height.Value, 15);
            Assert.Equal(50d, paragraph4.Image.Width.Value, 15);

            // or we can get any image and overwrite it's size
            document.Images[0].Height = 200;
            document.Images[0].Width = 200;

            Assert.Equal(200d, document.Images[0].Height.Value, 15);
            Assert.Equal(200d, document.Images[0].Width.Value, 15);

            var fileToSave = Path.Combine(_directoryDocuments, "CreatedDocumentWithImagesPrzemyslawKlysAndKulkozaurr.jpg");
            document.Images[0].SaveToFile(fileToSave);

            var fileInfo = new FileInfo(fileToSave);

            Assert.True(fileInfo.Length > 0);
            Assert.True(File.Exists(fileToSave));

            Assert.Equal(4, document.Images.Count);
            Assert.Equal(4, document.Sections[0].Images.Count);
            document.Save(false);
        }
    }
}
