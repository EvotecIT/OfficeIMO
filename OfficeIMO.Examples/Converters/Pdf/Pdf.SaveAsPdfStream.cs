using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfStreamRewind(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document and exporting to PDF stream");
            string docPath = Path.Combine(folderPath, "ExportToPdfStreamRewind.docx");
            string pdfPath = Path.Combine(folderPath, "ExportToPdfStreamRewind.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello World");
                document.Save();

                using (MemoryStream stream = new MemoryStream()) {
                    document.SaveAsPdf(stream);
                    using (FileStream fileStream = File.Create(pdfPath)) {
                        stream.CopyTo(fileStream);
                    }
                }
            }
        }
    }
}