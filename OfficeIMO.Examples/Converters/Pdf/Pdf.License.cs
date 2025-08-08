using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using QuestPDF.Infrastructure;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfWithLicense(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document and exporting to PDF with explicit QuestPDF license");
            string docPath = Path.Combine(folderPath, "PdfWithLicense.docx");
            string pdfPath = Path.Combine(folderPath, "PdfWithLicense.pdf");

            QuestPDF.Settings.License = null;

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello World");
                document.Save();

                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    QuestPdfLicenseType = LicenseType.Community
                });
            }

            if (openWord) {
                // openWord functionality not implemented
            }
        }
    }
}

