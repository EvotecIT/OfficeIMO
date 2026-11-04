using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageNumbers {
        internal static void Example_PageNumbers2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom page numbers 2");
            string filePath = System.IO.Path.Combine(folderPath, "Document with PageNumbers2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var table = document.PageSetup.Footer.AddTable(1, 2, WordTableStyle.TableGrid);
                table.Width = WordTableWidth.Percent(100);

                table.Rows[0].Cells[0].AddParagraph("Confidential");
                var para = table.Rows[0].Cells[1].AddParagraph();
                para.ParagraphAlignment = WordParagraphAlignment.Right;
                para.AddPageNumber(includeTotalPages: true, separator: " / ");

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
