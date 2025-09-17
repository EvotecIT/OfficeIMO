using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageNumbers {
        internal static void Example_PageNumbers5(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with section page numbers");
            string filePath = System.IO.Path.Combine(folderPath, "Document with PageNumbers5.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();

                var footers = Guard.NotNull(document.Footer, "Document footers must exist after enabling headers.");
                var defaultFooter = Guard.NotNull(footers.Default, "Default footer must exist after enabling headers.");

                var firstFooter = defaultFooter.AddParagraph();
                firstFooter.ParagraphAlignment = JustificationValues.Right;
                firstFooter.AddText("Page ");
                firstFooter.AddPageNumber(includeTotalPages: true, separator: " of ");

                document.AddParagraph("Section 1");

                var section = document.AddSection();
                section.AddPageNumbering(1);
                section.AddParagraph("Section 2");

                var secondFooter = defaultFooter.AddParagraph();
                secondFooter.ParagraphAlignment = JustificationValues.Right;
                secondFooter.AddText("Page ");
                secondFooter.AddPageNumber(includeTotalPages: true, separator: " of ");

                document.Save(openWord);
            }
        }
    }
}
