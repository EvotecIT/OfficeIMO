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
                var images = document.Images;
                Console.WriteLine("+ Document images: " + images.Count);

                var headers = Guard.NotNull(document.Header, "Document headers must exist when inspecting header images.");
                var defaultHeader = Guard.NotNull(headers.Default, "Default header must exist when inspecting header images.");
                var headerImages = defaultHeader.Images;
                Console.WriteLine("+ Document images in header: " + headerImages.Count);

                var footers = Guard.NotNull(document.Footer, "Document footers must exist when inspecting footer images.");
                var defaultFooter = Guard.NotNull(footers.Default, "Default footer must exist when inspecting footer images.");
                var footerImages = defaultFooter.Images;
                Console.WriteLine("+ Document images in footer: " + footerImages.Count);
                //document.Images[0].SaveToFile(System.IO.Path.Combine(outputPath, "random.jpg"));

                Console.WriteLine("----");
                foreach (var image in images) {
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
