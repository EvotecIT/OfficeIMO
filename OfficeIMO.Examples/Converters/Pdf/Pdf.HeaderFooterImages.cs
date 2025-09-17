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
                var headers = Guard.NotNull(document.Header, "Document headers must exist after enabling headers.");
                var defaultHeader = Guard.NotNull(headers.Default, "Default header must exist after enabling headers.");
                defaultHeader.AddParagraph().AddImage(imagePath, 50, 50);

                var footers = Guard.NotNull(document.Footer, "Document footers must exist after enabling headers.");
                var defaultFooter = Guard.NotNull(footers.Default, "Default footer must exist after enabling headers.");
                defaultFooter.AddParagraph().AddImage(imagePath, 300, 300);
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
