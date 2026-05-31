using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_PdfCustomFonts(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating PDF with a selected host font family");
            string docPath = Path.Combine(folderPath, "PdfCustomFonts.docx");
            string pdfPath = Path.Combine(folderPath, "PdfCustomFonts.pdf");
            string fontFamily = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? "Arial"
                : RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                    ? "Arial"
                    : "DejaVu Sans";

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("PDF paragraph using the selected host font family.");
                document.Save();

                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    FontFamily = fontFamily
                });
            }
        }
    }
}
