using OfficeIMO.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_HeaderFooterImages(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with header and footer images and exporting to PDF");
            string docPath = Path.Combine(folderPath, "PdfHeaderFooterImages.docx");
            string pdfPath = Path.Combine(folderPath, "PdfHeaderFooterImages.pdf");
            string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "EvotecLogo.png");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddHeadersAndFooters();
                document.Header.Default.AddParagraph().AddImage(imagePath, 50, 50);
                document.Footer.Default.AddParagraph().AddImage(imagePath, 300, 300);
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
