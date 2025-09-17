using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ReadWordWithImages() {
            Console.WriteLine("[*] Read Basic Word with Images");

            string outputPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            using (WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "BasicDocumentWithImages.docx"), true)) {
                Console.WriteLine("+ Document paragraphs: " + document.Paragraphs.Count);
                var images = document.Images;
                Console.WriteLine("+ Document images: " + images.Count);

                var firstImage = Guard.GetRequiredItem(images, 0, "Template should contain at least one image to export.");
                firstImage.SaveToFile(System.IO.Path.Combine(outputPath, "random.jpg"));
            }
        }
    }
}
