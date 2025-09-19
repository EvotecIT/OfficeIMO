using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ReadDocumentText {
        public static void Example_Pdf_ReadDocumentText(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.WithPageNumbers.pdf");
            if (!File.Exists(path)) WriterHeadersFooters.Example_Pdf_PageNumbers(folderPath, false);
            var doc = PdfReadDocument.Load(path, new PdfReadOptions { PreferToUnicode = true, UseWinAnsiFallback = true, AdjustKerningFromTJ = true });
            string text = doc.ExtractText();
            string outPath = Path.Combine(folderPath, "Pdf.WithPageNumbers.extracted.txt");
            File.WriteAllText(outPath, text);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = outPath, UseShellExecute = true });
        }
    }
}

