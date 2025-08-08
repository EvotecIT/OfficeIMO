using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static async Task Example_SaveAsPdfAsync(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document and exporting to PDF asynchronously");
            string docPath = Path.Combine(folderPath, "ExportToPdfAsync.docx");
            string pdfPath = Path.Combine(folderPath, "ExportToPdfAsync.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello Async PDF");
                document.Save();
                await document.SaveAsPdfAsync(pdfPath);
            }

            Console.WriteLine($"✓ Created: {pdfPath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
        }
    }
}

