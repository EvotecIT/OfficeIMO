using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageNumbers {
        internal static void Example_PageNumbers6(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with advanced page numbers");
            string filePath = System.IO.Path.Combine(folderPath, "Document with PageNumbers6.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].AddPageNumbering(1, NumberFormatValues.UpperRoman);
                document.AddHeadersAndFooters();

                var footers = Guard.NotNull(document.Footer, "Document footers must exist after enabling headers.");
                var defaultFooter = Guard.NotNull(footers.Default, "Default footer must exist after enabling headers.");

                var para = defaultFooter.AddParagraph();
                para.ParagraphAlignment = JustificationValues.Right;
                para.AddText("Page ");
                para.AddPageNumber(includeTotalPages: true, format: WordFieldFormat.Roman, separator: " of ");

                document.AddParagraph("First page");
                document.AddPageBreak();
                document.AddParagraph("Second page");

                document.Save(openWord);
            }
        }
    }
}
