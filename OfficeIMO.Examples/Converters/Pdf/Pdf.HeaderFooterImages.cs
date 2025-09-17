using OfficeIMO.Examples.Utils;
using OfficeIMO.Word.Pdf;
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
                var headers = Guard.NotNull(document.Header, "Headers should exist after calling AddHeadersAndFooters.");
                var defaultHeader = Guard.NotNull(headers.Default, "Default header should exist after calling AddHeadersAndFooters.");
                defaultHeader.AddParagraph().AddImage(imagePath, 50, 50);
                var footers = Guard.NotNull(document.Footer, "Footers should exist after calling AddHeadersAndFooters.");
                var defaultFooter = Guard.NotNull(footers.Default, "Default footer should exist after calling AddHeadersAndFooters.");
                defaultFooter.AddParagraph().AddImage(imagePath, 300, 300);
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
