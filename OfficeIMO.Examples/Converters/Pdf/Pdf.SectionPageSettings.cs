using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveSectionsPageSettings(string folderPath, bool openWord) {
            Console.WriteLine("[*] Exporting sections with different page settings to PDF");
            string docPath = Path.Combine(folderPath, "PdfSectionsPageSettings.docx");
            string pdfPath = Path.Combine(folderPath, "PdfSectionsPageSettings.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Sections[0].PageSettings.PageSize = WordPageSize.A4;
                document.Sections[0].PageSettings.Orientation = PageOrientationValues.Landscape;
                document.AddParagraph("Section 1");
                WordSection section2 = document.AddSection();
                section2.PageSettings.PageSize = WordPageSize.A5;
                section2.PageSettings.Orientation = PageOrientationValues.Portrait;
                section2.AddParagraph("Section 2");
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
            Console.WriteLine($"Created: {pdfPath}");
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
        }
    }
}
