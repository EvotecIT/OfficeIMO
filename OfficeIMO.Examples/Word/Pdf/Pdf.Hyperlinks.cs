using OfficeIMO.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfWithHyperlinks(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with hyperlinks and exporting to PDF");
            string docPath = Path.Combine(folderPath, "HyperlinksToPdf.docx");
            string pdfPath = Path.Combine(folderPath, "HyperlinksToPdf.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Visit ").AddHyperLink("OfficeIMO", new Uri("https://evotec.xyz"), addStyle: true);
                document.AddParagraph("Contact ").AddHyperLink("Email", new Uri("mailto:kontakt@evotec.pl"), addStyle: true);
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
