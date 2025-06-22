using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageNumbers {
        internal static void Example_PageNumbers3(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom page numbers 3");
            string filePath = System.IO.Path.Combine(folderPath, "Document with PageNumbers3.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].AddPageNumbering(2, NumberFormatValues.LowerRoman);
                document.AddHeadersAndFooters();

                var para = document.Footer.Default.AddParagraph();
                para.AddText("Page ");
                para.AddPageNumber();

                document.Save(openWord);
            }
        }
    }
}
