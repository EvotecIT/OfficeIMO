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


        [Fact]
        public void Test_CreatingWordDocumentWithImagesHeadersAndFooters() {
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            var filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithImagesHeadersAndFooters.docx");
            using var document = WordDocument.Create(filePath);

            var image = Path.Combine(_directoryWithImages, "PrzemyslawKlysAndKulkozaurr.jpg");

            document.AddHeadersAndFooters();
            document.DifferentOddAndEvenPages = true;

            Assert.True(document.Header.Default.Images.Count == 0);
            Assert.True(document.Header.Even.Images.Count == 0);
            Assert.True(document.Images.Count == 0);
            Assert.True(document.Footer.Default.Images.Count == 0);
            Assert.True(document.Footer.Even.Images.Count == 0);


            var header = document.Header.Default;
            // add image to header, directly to paragraph
            header.AddParagraph().AddImage(image, 100, 100);

            var footer = document.Footer.Default;
            // add image to footer, directly to paragraph
            footer.AddParagraph().AddImage(image, 200, 200);

            Assert.True(document.Header.Default.Images.Count == 1);
            Assert.True(document.Header.Even.Images.Count == 0);
            Assert.True(document.Images.Count == 0);
            Assert.True(document.Footer.Default.Images.Count == 1);
            Assert.True(document.Footer.Even.Images.Count == 0);

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);
            var paragraph = document.AddParagraph();
            using (var imageStream = System.IO.File.OpenRead(filePathImage)) {
                paragraph.AddImage(imageStream, fileNameImage, 300, 300);
            }
            Assert.True(document.Header.Default.Images.Count == 1);
            Assert.True(document.Header.Even.Images.Count == 0);

            Assert.True(document.Images.Count == 1);
            Assert.True(document.Footer.Default.Images.Count == 1);
            Assert.True(document.Footer.Even.Images.Count == 0);

            Assert.True(document.Images[0].FileName == fileNameImage);
            Assert.True(document.Images[0].Rotation == null);
            Assert.True(document.Images[0].Width == 300);
            Assert.True(document.Images[0].Height == 300);

            const string fileNameImageEvotec = "EvotecLogo.png";
            var filePathImageEvotec = System.IO.Path.Combine(imagePaths, fileNameImageEvotec);
            var paragraphHeader = document.Header.Even.AddParagraph();
            using (var imageStream = System.IO.File.OpenRead(filePathImageEvotec)) {
                paragraphHeader.AddImage(imageStream, fileNameImageEvotec, 300, 300, WrapImageText.InLineWithText, "This is a test");

            }

            Assert.True(document.Header.Default.Images.Count == 1);
            Assert.True(document.Header.Even.Images.Count == 1);
            Assert.True(document.Header.Even.Images[0].FileName == fileNameImageEvotec);
            Assert.True(document.Header.Even.Images[0].Description == "This is a test");

            document.Header.Even.Images[0].Description = "Different description";
            Assert.True(document.Header.Even.Images[0].VerticalFlip == null);
            Assert.True(document.Header.Even.Images[0].HorizontalFlip == null);
            document.Header.Even.Images[0].VerticalFlip = true;

            Assert.True(document.Header.Even.Images[0].Description == "Different description");
            Assert.True(document.Header.Even.Images[0].VerticalFlip == true);
            document.Save();
        }
    }

}
