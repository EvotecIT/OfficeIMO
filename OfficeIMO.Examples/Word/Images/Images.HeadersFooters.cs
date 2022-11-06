using System;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_AddingImagesHeadersFooters(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some Images");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithImagesHeaderFooters.docx");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is sparta";
                document.BuiltinDocumentProperties.Creator = "Przemek";
                var filePathImage = System.IO.Path.Combine(imagePaths, "Kulek.jpg");

                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;

                var header = document.Header.Default;
                var paragraphHeader = header.AddParagraph("This is header");

                // add image to header, directly to paragraph
                header.AddParagraph().AddImage(filePathImage, 100, 100);

                // add image to footer, directly to paragraph
                document.Footer.Default.AddParagraph().AddImage(filePathImage, 100, 100);

                // add image to header, but to a table
                var table = header.AddTable(2, 2);
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Test123";
                table.Rows[1].Cells[0].Paragraphs[0].AddImage(filePathImage, 50, 50);
                table.Alignment = TableRowAlignmentValues.Right;

                var paragraph = document.AddParagraph("This paragraph starts with some text");
                paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";
                paragraph.Bold = true;

                // add table with an image, but to document
                var table1 = document.AddTable(2, 2);
                table1.Rows[1].Cells[1].Paragraphs[0].Text = "Test - In document";
                table1.Rows[1].Cells[0].Paragraphs[0].AddImage(filePathImage, 50, 50);


                // lets add image to paragraph
                paragraph.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22);
                //paragraph.Image.WrapText = true; // WrapSideValues.Both;

                var paragraph5 = paragraph.AddText("and more text");
                paragraph5.Bold = true;

                document.AddParagraph("This adds another picture with 500x500");


                WordParagraph paragraph2 = document.AddParagraph();
                paragraph2.AddImage(filePathImage, 500, 500);
                //paragraph2.Image.BlackWiteMode = BlackWhiteModeValues.GrayWhite;
                paragraph2.Image.Rotation = 180;
                paragraph2.Image.Shape = ShapeTypeValues.ActionButtonMovie;


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

                // or we can get any image and overwrite it's size
                document.Images[0].Height = 200;
                document.Images[0].Width = 200;

                string fileToSave = System.IO.Path.Combine(imagePaths, "OutputPrzemyslawKlysAndKulkozaurr.jpg");
                document.Images[0].SaveToFile(fileToSave);

                var paragraphHeaderEven = document.Header.Even.AddParagraph("This adds another picture via Stream with 100x100 to Header Even");
                const string fileNameImageEvotec = "EvotecLogo.png";
                var filePathImageEvotec = System.IO.Path.Combine(imagePaths, fileNameImageEvotec);
                using (var imageStream = System.IO.File.OpenRead(filePathImageEvotec)) {
                    paragraphHeaderEven.AddImage(imageStream, fileNameImageEvotec, 100, 100);
                }

                //var filePathImageEvotecSave = System.IO.Path.Combine(imagePaths, "savedFile.png");
                //paragraphHeaderEven.Image.SaveToFile(filePathImageEvotecSave);

                document.Save(openWord);
            }
        }
    }
}
