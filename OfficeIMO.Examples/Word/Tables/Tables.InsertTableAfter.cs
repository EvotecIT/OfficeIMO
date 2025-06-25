using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_InsertTableAfterSimple(string folderPath, bool openWord) {
            Console.WriteLine("[*] Inserting table after a paragraph");
            string filePath = Path.Combine(folderPath, "Example-InsertTableAfter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var before = document.AddParagraph("Before");
                document.AddParagraph("After");

                var table = document.CreateTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "T1";

                document.InsertTableAfter(before, table);
                document.Save(openWord);
            }
        }

        internal static void Example_InsertTableAfterAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Inserting table with style and additional paragraph");
            string filePath = Path.Combine(folderPath, "Example-InsertTableAfterAdvanced.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var start = document.AddParagraph("Start");
                document.AddParagraph("End");

                // Insert a paragraph in between
                document.InsertParagraphAt(1).Text = "Middle";

                var table = document.CreateTable(3, 3, WordTableStyle.GridTable5DarkAccent1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Advanced";

                document.InsertTableAfter(start, table);
                document.Save(openWord);
            }
        }
    }
}
