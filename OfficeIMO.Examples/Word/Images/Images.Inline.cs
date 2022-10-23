using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_AddingImagesInline(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with inline images");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithInlineImages2.docx");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var file = System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg");
                var paragraph = document.AddParagraph();
                var pargraphWithImage = paragraph.AddImage(file, 100, 100, "Przemek and Kulek on an image");

                Console.WriteLine("Image is inline: " + pargraphWithImage.Image.Rotation);

                pargraphWithImage.Image.VerticalFlip = false;
                pargraphWithImage.Image.HorizontalFlip = false;
                pargraphWithImage.Image.Rotation = 190;
                pargraphWithImage.Image.Shape = ShapeTypeValues.Cloud;
                pargraphWithImage.Image.BlackWiteMode = BlackWhiteModeValues.GrayWhite;
                pargraphWithImage.Image.Description = "Other description";

                Console.WriteLine("Image is inline: " + pargraphWithImage.Image.Rotation);

                document.Save(openWord);
            }
        }
    }
}
