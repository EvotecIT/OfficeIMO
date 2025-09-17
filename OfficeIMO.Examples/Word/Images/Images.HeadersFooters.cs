using System;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
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

                var header = GetDocumentHeaderOrThrow(document);
                var paragraphHeader = header.AddParagraph("This is header");

                // add image to header, directly to paragraph
                header.AddParagraph().AddImage(filePathImage, 100, 100);

                // add image to footer, directly to paragraph
                var footer = GetDocumentFooterOrThrow(document);
                footer.AddParagraph().AddImage(filePathImage, 100, 100);

                // add image to header, but to a table
                var table = header.AddTable(2, 2);
                GetOrAddParagraph(table, 1, 1).Text = "Test123";
                GetOrAddParagraph(table, 1, 0).AddImage(filePathImage, 50, 50);
                table.Alignment = TableRowAlignmentValues.Right;

                var paragraph = document.AddParagraph("This paragraph starts with some text");
                paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";
                paragraph.Bold = true;

                // add table with an image, but to document
                var table1 = document.AddTable(2, 2);
                GetOrAddParagraph(table1, 1, 1).Text = "Test - In document";
                GetOrAddParagraph(table1, 1, 0).AddImage(filePathImage, 50, 50);


                // lets add image to paragraph
                paragraph.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22);
                //paragraph.Image.WrapText = true; // WrapSideValues.Both;

                var paragraph5 = paragraph.AddText("and more text");
                paragraph5.Bold = true;

                document.AddParagraph("This adds another picture with 500x500");


                WordParagraph paragraph2 = document.AddParagraph();
                paragraph2.AddImage(filePathImage, 500, 500);
                var paragraph2Image = Guard.NotNull(paragraph2.Image, "Paragraph should contain the newly added image.");
                //paragraph2Image.BlackWiteMode = BlackWhiteModeValues.GrayWhite;
                paragraph2Image.Rotation = 180;
                paragraph2Image.Shape = ShapeTypeValues.ActionButtonMovie;


                document.AddParagraph("This adds another picture with 100x100");

                WordParagraph paragraph3 = document.AddParagraph();
                paragraph3.AddImage(filePathImage, 100, 100);

                // we add paragraph with an image
                WordParagraph paragraph4 = document.AddParagraph();
                paragraph4.AddImage(filePathImage);

                var paragraph4Image = Guard.NotNull(paragraph4.Image, "Paragraph should contain the added image.");

                // we can get the height of the image from paragraph
                Console.WriteLine("This document has image, which has height of: " + paragraph4Image.Height + " pixels (I think) ;-)");

                // we can also overwrite height later on
                paragraph4Image.Height = 50;
                paragraph4Image.Width = 50;
                // this doesn't work
                paragraph4Image.HorizontalFlip = true;

                // or we can get any image and overwrite it's size
                var firstImage = Guard.GetRequiredItem(document.Images, 0, "Document should contain at least one image.");
                firstImage.Height = 200;
                firstImage.Width = 200;

                string fileToSave = System.IO.Path.Combine(imagePaths, "OutputPrzemyslawKlysAndKulkozaurr.jpg");
                firstImage.SaveToFile(fileToSave);

                var headerEven = GetDocumentHeaderOrThrow(document, HeaderFooterValues.Even);
                var paragraphHeaderEven = headerEven.AddParagraph("This adds another picture via Stream with 100x100 to Header Even");
                const string fileNameImageEvotec = "EvotecLogo.png";
                var filePathImageEvotec = System.IO.Path.Combine(imagePaths, fileNameImageEvotec);
                using (var imageStream = System.IO.File.OpenRead(filePathImageEvotec)) {
                    paragraphHeaderEven.AddImage(imageStream, fileNameImageEvotec, 100, 100);
                }

                //var filePathImageEvotecSave = System.IO.Path.Combine(imagePaths, "savedFile.png");
                //paragraphHeaderEven.Image.SaveToFile(filePathImageEvotecSave);

                document.Save(openWord);

                static WordParagraph GetOrAddParagraph(WordTable table, int rowIndex, int columnIndex) {
                    var row = Guard.GetRequiredItem(table.Rows, rowIndex, $"Table must contain row index {rowIndex}.");
                    var cell = Guard.GetRequiredItem(row.Cells, columnIndex, $"Row must contain cell index {columnIndex}.");
                    return cell.Paragraphs.FirstOrDefault() ?? cell.AddParagraph();
                }
            }
        }
    }
}
