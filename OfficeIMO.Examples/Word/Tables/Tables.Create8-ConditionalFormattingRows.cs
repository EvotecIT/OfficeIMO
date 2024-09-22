using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_BasicTables8(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tables and conditional formatting");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables8.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.AddParagraph();

                WordTable wordTable = document.AddTable(3, 5, WordTableStyle.PlainTable1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[0].Cells[1].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[1].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[1].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[0].Cells[2].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[2].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[2].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[0].Cells[3].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[3].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[3].Paragraphs[0].Text = "Test 3";

                wordTable.ConditionalFormattingFirstRow = true;
                wordTable.ConditionalFormattingLastRow = true;
                wordTable.ConditionalFormattingFirstColumn = false;

                document.Save(openWord);
            }
        }
    }
}
