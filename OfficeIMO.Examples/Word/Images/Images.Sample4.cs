using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_AddingImagesSample4(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some Images and Samples");
            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithImagesSample4.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);

            var paragraph1 = document.AddParagraph("This paragraph starts with some text");
            paragraph1.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200);
            paragraph1.Image.Shape = ShapeTypeValues.Cube;

            var paragraph2 = document.AddParagraph("Image will be placed behind text");
            paragraph2.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapTextImage.BehindText, "Przemek and Kulek on an image");


            var paragraph3 = document.AddParagraph("Image will be in front of text");
            paragraph3.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapTextImage.InFrontOfText, "Przemek and Kulek on an image");


            var paragraph5 = document.AddParagraph("Image will be Square");
            paragraph5.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapTextImage.Square, "Przemek and Kulek on an image");


            var paragraph6 = document.AddParagraph("Image will be Through");
            paragraph6.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapTextImage.Through, "Przemek and Kulek on an image");


            var paragraph7 = document.AddParagraph("Image will be Tight");
            paragraph7.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapTextImage.Tight, "Przemek and Kulek on an image");


            var paragraph8 = document.AddParagraph("Image will be Top And Bottom");
            paragraph8.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapTextImage.TopAndBottom, "Przemek and Kulek on an image");
            paragraph8.Image.Shape = ShapeTypeValues.Can;

            document.Save(openWord);
        }
    }
}
