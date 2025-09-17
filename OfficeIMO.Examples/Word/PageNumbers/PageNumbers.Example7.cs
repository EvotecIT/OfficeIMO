using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageNumbers {
        internal static void Example_PageNumbers7(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom page numbers 7");
            string filePath = System.IO.Path.Combine(folderPath, "Document with PageNumbers7.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var pageNumber = GetDocumentFooterOrThrow(document).AddPageNumber(WordPageNumberStyle.PlainNumber);
                pageNumber.AppendText(" of ");
                pageNumber.Paragraph.AddField(WordFieldType.NumPages);
                document.Save(openWord);
            }
        }
    }
}
