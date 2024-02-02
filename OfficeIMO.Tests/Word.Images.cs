using System;
using System.IO;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
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
            Assert.True(paragraph.Image.WrapText == WrapTextImage.InLineWithText);

            paragraph.Image.WrapText = WrapTextImage.BehindText;

            Assert.True(paragraph.Image.WrapText == WrapTextImage.BehindText);

            paragraph.Image.WrapText = WrapTextImage.InLineWithText;

            Assert.True(paragraph.Image.WrapText == WrapTextImage.InLineWithText);

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

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
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
                paragraphHeader.AddImage(imageStream, fileNameImageEvotec, 300, 300, WrapTextImage.InLineWithText, "This is a test");
                Assert.True(paragraphHeader.Image.CompressionQuality == BlipCompressionValues.Print);
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
            Assert.True(document.Header.Even.Images[0].CompressionQuality == BlipCompressionValues.Print);
            document.Header.Even.Images[0].CompressionQuality = BlipCompressionValues.HighQualityPrint;
            Assert.True(document.Header.Even.Images[0].CompressionQuality == BlipCompressionValues.HighQualityPrint);

            document.Save();
        }



        [Fact]
        public void Test_LoadingWordDocumentWithImages() {
            var documentsPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            var filePath = Path.Combine(documentsPaths, "DocumentWithImagesWraps.docx");
            using (var document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs.Count == 36);
                Assert.True(document.Images.Count == 4);
                Assert.True(document.Header.Default.Images.Count == 1);
                Assert.True(document.Footer.Default.Images.Count == 0);

                Assert.True(document.Images[0].WrapText == WrapTextImage.InLineWithText);
                Assert.True(document.Images[1].WrapText == WrapTextImage.Square);
                Assert.True(document.Images[2].WrapText == WrapTextImage.InFrontOfText);
                Assert.True(document.Images[3].WrapText == WrapTextImage.BehindText);
                Assert.True(document.Header.Default.Images[0].WrapText == WrapTextImage.InLineWithText);

                Assert.True(document.Images[0].Shape == ShapeTypeValues.Rectangle);
                Assert.True(document.Images[1].Shape == ShapeTypeValues.Rectangle);
                Assert.True(document.Images[2].Shape == ShapeTypeValues.Rectangle);
                Assert.True(document.Images[3].Shape == ShapeTypeValues.Rectangle);

                document.Images[0].Shape = ShapeTypeValues.Cloud;

                Assert.True(document.Images[0].Shape == ShapeTypeValues.Cloud);
                document.Save(false);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithImagesWraps() {
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImagesWraps.docx");
            using var document = WordDocument.Create(filePath);

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);

            var paragraph1 = document.AddParagraph("This is a test document with images wraps");
            paragraph1.AddImage(filePathImage, 100, 100, WrapTextImage.InLineWithText);

            var paragraph2 = document.AddParagraph("This is a test document with images wraps");
            paragraph2.AddImage(filePathImage, 100, 100, WrapTextImage.BehindText);

            var paragraph3 = document.AddParagraph("This is a test document with images wraps");
            paragraph3.AddImage(filePathImage, 100, 100, WrapTextImage.InFrontOfText);

            var paragraph4 = document.AddParagraph("This is a test document with images wraps");
            paragraph4.AddImage(filePathImage, 100, 100, WrapTextImage.TopAndBottom);

            var paragraph5 = document.AddParagraph("This is a test document with images wraps");
            paragraph5.AddImage(filePathImage, 100, 100, WrapTextImage.Square);

            var paragraph6 = document.AddParagraph("This is a test document with images wraps");
            paragraph6.AddImage(filePathImage, 100, 100, WrapTextImage.Tight);

            var paragraph7 = document.AddParagraph("This is a test document with images wraps");
            paragraph7.AddImage(filePathImage, 100, 100, WrapTextImage.Through);

            Assert.True(document.Paragraphs.Count == 7);
            Assert.True(document.Paragraphs[0].Image.WrapText == WrapTextImage.InLineWithText);
            Assert.True(document.Paragraphs[1].Image.WrapText == WrapTextImage.BehindText);
            Assert.True(document.Paragraphs[2].Image.WrapText == WrapTextImage.InFrontOfText);
            Assert.True(document.Paragraphs[3].Image.WrapText == WrapTextImage.TopAndBottom);
            Assert.True(document.Paragraphs[4].Image.WrapText == WrapTextImage.Square);
            Assert.True(document.Paragraphs[5].Image.WrapText == WrapTextImage.Tight);
            Assert.True(document.Paragraphs[6].Image.WrapText == WrapTextImage.Through);

            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }

        [Fact]
        public void Test_CreatingWordDocumentWithFixedImages() {
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            var filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithFixedImages.docx");
            using var document = WordDocument.Create(filePath);

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);

            var paragraph1 = document.AddParagraph("This is a test document with images wraps");
            paragraph1.AddImage(filePathImage, 100, 100, WrapTextImage.InLineWithText);

            var paragraph2 = document.AddParagraph("This is a test document with images wraps");
            paragraph2.AddImage(filePathImage, 100, 100, WrapTextImage.Square);

            Assert.True(document.Paragraphs.Count == 2);
            Assert.True(document.Paragraphs[0].Image.WrapText == WrapTextImage.InLineWithText);
            Assert.True(document.Paragraphs[1].Image.WrapText == WrapTextImage.Square);
            Assert.Throws<System.InvalidOperationException>(() => document.Paragraphs[0].Image.horizontalPosition);
            Assert.Throws<System.InvalidOperationException>(() => document.Paragraphs[0].Image.verticalPosition);

            int emus = 914400;
            HorizontalRelativePositionValues hRelativeFromGood = HorizontalRelativePositionValues.RightMargin;
            HorizontalRelativePositionValues hRelativeFromFail = HorizontalRelativePositionValues.LeftMargin;
            VerticalRelativePositionValues vRelativeFromGood = VerticalRelativePositionValues.Page;
            VerticalRelativePositionValues vRelativeFromFail = VerticalRelativePositionValues.Line;

            HorizontalPosition horizontalPosition1 = new HorizontalPosition() {
                RelativeFrom = hRelativeFromGood,
                PositionOffset = new PositionOffset { Text = $"{emus}" }
            };

            VerticalPosition verticalPosition1 = new VerticalPosition() {
                RelativeFrom = vRelativeFromGood,
                PositionOffset = new PositionOffset { Text = $"{emus}" }
            };

            Assert.Throws<System.InvalidOperationException>(() => document.Paragraphs[0].Image.horizontalPosition = horizontalPosition1);
            Assert.Throws<System.InvalidOperationException>(() => document.Paragraphs[0].Image.verticalPosition = verticalPosition1);

            PositionOffset positionOffsetGood = new PositionOffset { Text = $"{emus}" };
            PositionOffset positionOffsetFail = new PositionOffset { Text = $"{2 * emus}" };
            Assert.NotEqual(positionOffsetFail.Text, positionOffsetGood.Text);

            document.Paragraphs[1].Image.horizontalPosition = horizontalPosition1;
            Assert.Equal(positionOffsetGood.Text, document.Paragraphs[1].Image.horizontalPosition.PositionOffset.Text);
            Assert.NotEqual(positionOffsetFail.Text, document.Paragraphs[1].Image.horizontalPosition.PositionOffset.Text);
            Assert.Equal(hRelativeFromGood, document.Paragraphs[1].Image.horizontalPosition.RelativeFrom.Value);
            Assert.NotEqual(hRelativeFromFail, document.Paragraphs[1].Image.horizontalPosition.RelativeFrom.Value);


            document.Paragraphs[1].Image.verticalPosition = verticalPosition1;
            Assert.Equal(positionOffsetGood.Text, document.Paragraphs[1].Image.verticalPosition.PositionOffset.Text);
            Assert.NotEqual(positionOffsetFail.Text, document.Paragraphs[1].Image.verticalPosition.PositionOffset.Text);
            Assert.Equal(vRelativeFromGood, document.Paragraphs[1].Image.verticalPosition.RelativeFrom.Value);
            Assert.NotEqual(vRelativeFromFail, document.Paragraphs[1].Image.verticalPosition.RelativeFrom.Value);

            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }

        [Fact]
        public void Test_CreatingWordDocumentWithImagesInTable() {
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImagesInTable.docx");
            using var document = WordDocument.Create(filePath);

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);

            var table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].AddImage(filePathImage, 200, 200);

            // not really necessary to add new paragraph since one is already there by default
            var paragraph = table.Rows[0].Cells[1].AddParagraph();
            paragraph.AddImage(filePathImage, 200, 200);

            document.AddHeadersAndFooters();

            var tableInHeader = document.Header.Default.AddTable(2, 2);
            tableInHeader.Rows[0].Cells[0].Paragraphs[0].AddImage(filePathImage, 200, 200);

            // not really necessary to add new paragraph since one is already there by default
            var paragraphInHeader = tableInHeader.Rows[0].Cells[1].AddParagraph();
            paragraphInHeader.AddImage(filePathImage, 200, 200);

            Assert.True(document.Tables.Count == 1);
            Assert.True(document.Tables[0].Rows.Count == 2);
            Assert.True(document.Tables[0].Rows[0].Cells.Count == 2);

            Assert.True(document.Header.Default.Tables.Count == 1);

            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }

        [Fact]
        public void Test_CreatingWordDocumentWithImagesInline() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImagesInline.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph("This paragraph starts with some text");
            paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

            // lets add image to paragraph
            var imageParagraph = paragraph.AddImage(Path.Combine(_directoryWithImages, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22, WrapTextImage.InLineWithText);

            Assert.True(document.Images[0].WrapText == WrapTextImage.InLineWithText);

            Assert.True(imageParagraph.Image.WrapText == WrapTextImage.InLineWithText);

            imageParagraph.Image.WrapText = WrapTextImage.Square;

            Assert.True(imageParagraph.Image.WrapText == WrapTextImage.Square);

            var paragraph5 = paragraph.AddText("and more text");
            paragraph5.Bold = true;


            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }

    }

}
