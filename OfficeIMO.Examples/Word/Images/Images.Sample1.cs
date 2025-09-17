using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_AddingImages(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some Images");
            //string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            //string imagePaths = System.IO.Path.Combine(baseDirectory, "Images");
            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithImages.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "This is sparta";
            document.BuiltinDocumentProperties.Creator = "Przemek";

            var paragraph1 = document.AddParagraph("This paragraph starts with some text");
            paragraph1.Text = "0th This paragraph started with some other text and was overwritten and made bold.";
            // lets add image to paragraph
            var paragraphImage = paragraph1.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22, WrapTextImage.BehindText);

            var paragraph1Image = Guard.NotNull(paragraph1.Image, "Paragraph should contain the inserted image.");
            var paragraphImageImage = Guard.NotNull(paragraphImage.Image, "Paragraph should contain the inserted image.");

            Console.WriteLine(paragraph1Image.WrapText);
            Console.WriteLine(paragraphImageImage.WrapText);

            var paragraph2 = paragraph1.AddText("and more text");
            paragraph2.Bold = true;

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);

            document.AddParagraph("This adds another picture with 500x500");
            var paragraph3 = document.AddParagraph();
            paragraph3.AddImage(filePathImage, 500, 500);
            var paragraph3Image = Guard.NotNull(paragraph3.Image, "Paragraph should contain the large image.");
            //paragraph3Image.BlackWiteMode = BlackWhiteModeValues.GrayWhite;
            paragraph3Image.Rotation = 180;
            paragraph3Image.Shape = ShapeTypeValues.ActionButtonMovie;

            document.AddParagraph("This adds another picture with 100x100");
            var paragraph4 = document.AddParagraph();
            paragraph4.AddImage(filePathImage, 100, 100);

            document.AddParagraph("This adds another picture via Stream with 100x100");
            var paragraph5 = document.AddParagraph();
            using (var imageStream = System.IO.File.OpenRead(filePathImage)) {
                paragraph5.AddImage(imageStream, fileNameImage, 100, 100);
            }

            // we add paragraph with an image
            var paragraph6 = document.AddParagraph();
            paragraph6.AddImage(filePathImage);
            var paragraph6Image = Guard.NotNull(paragraph6.Image, "Paragraph should contain the default-sized image.");
            // we can get the height of the image from paragraph
            Console.WriteLine("This document has image, which has height of: " + paragraph6Image.Height + " pixels (I think) ;-)");
            // we can also overwrite height later on
            paragraph6Image.Height = 50;
            paragraph6Image.Width = 50;
            // this doesn't work
            paragraph6Image.HorizontalFlip = true;

            // or we can get any image and overwrite it's size
            var firstImage = Guard.GetRequiredItem(document.Images, 0, "Document should contain at least one image.");
            firstImage.Height = 200;
            firstImage.Width = 200;

            var fileToSave = System.IO.Path.Combine(imagePaths, "OutputPrzemyslawKlysAndKulkozaurr.jpg");
            firstImage.SaveToFile(fileToSave);

            document.Save(openWord);
        }
    }
}
