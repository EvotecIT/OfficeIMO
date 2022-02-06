using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = System.Drawing.Color;
using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_CreatingWordDocumentWithImages() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImages.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is sparta";
                document.BuiltinDocumentProperties.Creator = "Przemek";

                var paragraph = document.AddParagraph("This paragraph starts with some text");
                paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

                // lets add image to paragraph
                paragraph.AddImage(System.IO.Path.Combine(_directoryWithImages, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22);
                //paragraph.Image.WrapText = true; // WrapSideValues.Both;

                var paragraph5 = paragraph.AddText("and more text");
                paragraph5.Bold = true;

                Assert.True(paragraph.Image.Width == 22);
                Assert.True(paragraph.Image.Width == 22);

                document.AddParagraph("This adds another picture with 500x500");

                var filePathImage = System.IO.Path.Combine(_directoryWithImages, "Kulek.jpg");
                WordParagraph paragraph2 = document.AddParagraph();
                paragraph2.AddImage(filePathImage, 500, 500);
                //paragraph2.Image.BlackWiteMode = BlackWhiteModeValues.GrayWhite;
                paragraph2.Image.Rotation = 180;
                paragraph2.Image.Shape = ShapeTypeValues.ActionButtonMovie;

                Assert.True(paragraph2.Image.Height == 500);
                Assert.True(paragraph2.Image.Width == 500);

                document.AddParagraph("This adds another picture with 100x100");

                WordParagraph paragraph3 = document.AddParagraph();
                paragraph3.AddImage(filePathImage, 100, 100);

                // we add paragraph with an image
                WordParagraph paragraph4 = document.AddParagraph();
                paragraph4.AddImage(filePathImage);

                // we can get the height of the image from paragraph
                Console.WriteLine("This document has image, which has height of: " + paragraph4.Image.Height + " pixels (I think) ;-)");

                // we can also overwrite height later on
                paragraph4.Image.Height = 50;
                paragraph4.Image.Width = 50;
                // this doesn't work
                paragraph4.Image.HorizontalFlip = true;


                Assert.True(paragraph4.Image.Height == 50);
                Assert.True(paragraph4.Image.Width == 50);

                // or we can get any image and overwrite it's size
                document.ImagesList[0].Height = 200;
                document.ImagesList[0].Width = 200;

                Assert.True(document.ImagesList[0].Height == 200);
                Assert.True(document.ImagesList[0].Height == 200);

                string fileToSave = System.IO.Path.Combine(_directoryDocuments, "CreatedDocumentWithImagesPrzemyslawKlysAndKulkozaurr.jpg");
                document.ImagesList[0].SaveToFile(fileToSave);

                var fileInfo = new FileInfo(fileToSave);

                Assert.True(fileInfo.Length > 0);
                Assert.True(File.Exists(fileToSave) == true);

                Assert.True(document.Images.Count == 4);
                Assert.True(document.Sections[0].Images.Count == 4);
                document.Save(false);
            }
        }
    }
}
