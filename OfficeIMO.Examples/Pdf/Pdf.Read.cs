using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ReadPdf {
        public static void Example_Pdf_ReadPlainText(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "HelloWorld.OfficeIMO.Pdf.pdf");
            if (!File.Exists(path)) return; // run Basic first
            string text = PdfTextExtractor.ExtractAllText(path);
            string outPath = Path.Combine(folderPath, "HelloWorld.OfficeIMO.Pdf.extracted.txt");
            File.WriteAllText(outPath, text);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = outPath, UseShellExecute = true });
        }
    }
}

