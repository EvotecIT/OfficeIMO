using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static async Task Example_PdfInterfaceAsync(string folderPath, bool openWord) {
            Console.WriteLine("[*] Exporting to PDF via interface async");
            string pdfPath = Path.Combine(folderPath, "ExportInterfaceAsync.pdf");
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Hello PDF");
            using MemoryStream docStream = new MemoryStream();
            document.Save(docStream);
            docStream.Position = 0;
            using MemoryStream pdfStream = new MemoryStream();
            ConverterRegistry.Register("word->pdf", () => new WordPdfConverter());
            IWordConverter converter = ConverterRegistry.Resolve("word->pdf");
            await converter.ConvertAsync(docStream, pdfStream, new PdfSaveOptions());
            await File.WriteAllBytesAsync(pdfPath, pdfStream.ToArray());
        }
    }
}
