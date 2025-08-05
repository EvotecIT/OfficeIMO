using System;
using System.IO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_PdfInterface(string folderPath, bool openWord) {
            Console.WriteLine("[*] Exporting to PDF via interface");
            string pdfPath = Path.Combine(folderPath, "ExportInterface.pdf");
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Hello PDF");
            using MemoryStream docStream = new MemoryStream();
            document.Save(docStream);
            docStream.Position = 0;
            using MemoryStream pdfStream = new MemoryStream();
            ConverterRegistry.Register("word->pdf", () => new WordPdfConverter());
            IWordConverter converter = ConverterRegistry.Resolve("word->pdf");
            converter.Convert(docStream, pdfStream, new PdfSaveOptions());
            File.WriteAllBytes(pdfPath, pdfStream.ToArray());
        }
    }
}
