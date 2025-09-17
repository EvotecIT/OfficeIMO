using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageNumbers {
        internal static void Example_PageNumbers8(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating documents with various custom page number formats");

            string[] formats = new[] {
                "0", "00", "000", "0000", "#", "##", "###", "#,##0", "0.00", "##0.##",
                "000#", "#000", "10-20", "Page 0", "0-00"
            };
            foreach (var format in formats) {
                string safeFormat = System.Text.RegularExpressions.Regex.Replace(format, "[^A-Za-z0-9]", "_");
                string filePath = System.IO.Path.Combine(folderPath, $"Document_PageNumbers_{safeFormat}.docx");
                using (WordDocument document = WordDocument.Create(filePath)) {
                    document.AddHeadersAndFooters();
                    var pageNumber = GetDocumentFooterOrThrow(document).AddPageNumber(WordPageNumberStyle.PlainNumber);
                    pageNumber.CustomFormat = format;
                    document.Save(openWord);
                }
            }
        }
    }
}
