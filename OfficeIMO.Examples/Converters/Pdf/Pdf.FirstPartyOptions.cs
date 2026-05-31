using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfWithFirstPartyOptions(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document and exporting to PDF with first-party OfficeIMO options");
            string docPath = Path.Combine(folderPath, "PdfWithFirstPartyOptions.docx");
            string pdfPath = Path.Combine(folderPath, "PdfWithFirstPartyOptions.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello World");
                document.Save();

                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    OfficeIMOPageSize = OfficeIMO.Pdf.PageSizes.Letter,
                    OfficeIMOMargins = OfficeIMO.Pdf.PageMargins.Narrow
                });
            }

            if (openWord) {
                // openWord functionality not implemented
            }
        }
    }
}

