using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageNumbers {
        internal static void Example_PageNumbers4(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom page numbers 4");
            string filePath = System.IO.Path.Combine(folderPath, "Document with PageNumbers4.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();

                var headers = Guard.NotNull(document.Header, "Document headers must exist after enabling headers.");
                var defaultHeader = Guard.NotNull(headers.Default, "Default header must exist after enabling headers.");

                var para = defaultHeader.AddParagraph();
                para.ParagraphAlignment = JustificationValues.Center;
                para.AddPageNumber();

                document.Save(openWord);
            }
        }
    }
}
