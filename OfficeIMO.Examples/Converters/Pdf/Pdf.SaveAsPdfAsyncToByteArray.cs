using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static async Task Example_SaveAsPdfAsyncToByteArray(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document and exporting to PDF asynchronously as bytes");
            string docPath = Path.Combine(folderPath, "ExportToPdfAsyncBytes.docx");
            string pdfPath = Path.Combine(folderPath, "ExportToPdfAsyncBytes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello Async PDF Bytes");
                document.Save();
                byte[] pdfBytes = await document.SaveAsPdfAsync(cancellationToken: CancellationToken.None);
                File.WriteAllBytes(pdfPath, pdfBytes);
            }

            Console.WriteLine($"✓ Created: {pdfPath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
        }
    }
}
