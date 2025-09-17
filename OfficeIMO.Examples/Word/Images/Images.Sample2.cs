using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ReadWordWithImagesAndDiffWraps() {
            Console.WriteLine("[*] Read Basic Word with Images and different wraps");

            string outputPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            using (WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "DocumentWithImagesWraps.docx"), true)) {
                Console.WriteLine("+ Document paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ Document images: " + document.Images.Count);
                var headers = Guard.NotNull(document.Header, "Loaded document is expected to contain headers.");
                var headerDefault = Guard.NotNull(headers.Default, "Loaded document must contain a default header.");
                Console.WriteLine("+ Document images in header: " + headerDefault.Images.Count);
                var footers = Guard.NotNull(document.Footer, "Loaded document is expected to contain footers.");
                var footerDefault = Guard.NotNull(footers.Default, "Loaded document must contain a default footer.");
                Console.WriteLine("+ Document images in footer: " + footerDefault.Images.Count);
                //document.Images[0].SaveToFile(System.IO.Path.Combine(outputPath, "random.jpg"));

                Console.WriteLine("----");
                foreach (var image in document.Images) {
                    Console.WriteLine("+ Image: " + image.FileName);
                    Console.WriteLine("+ Image: " + image.Width);
                    Console.WriteLine("+ Image: " + image.Height);
                    Console.WriteLine("+ Image: " + image.WrapText);
                }
                Console.WriteLine("----");
            }
        }
    }
}
