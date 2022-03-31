using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_BasicTablesLoad1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading standard document with tables");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables1.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.GridTable1LightAccent5);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                wordTable = document.AddTable(3, 4, WordTableStyle.GridTable1LightAccent6);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                wordTable = document.AddTable(3, 4, WordTableStyle.GridTable1LightAccent3);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                WordTable wordTableFromEarlier = document.Tables[0];
                wordTableFromEarlier.Rows[1].Cells[1].Paragraphs[0].Text = "Middle table";

                document.Save(openWord);
            }
        }
    }
}
