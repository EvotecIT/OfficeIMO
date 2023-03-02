using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_BasicTables8_StylesModification(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tables");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables8_StyleModification.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.AddParagraph();

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                Console.WriteLine("Table style: " + wordTable.Style);
                Console.WriteLine("Table MarginDefaultTopWidth: " + wordTable.StyleDetails.MarginDefaultTopWidth);

                wordTable.Style = WordTableStyle.GridTable1Light;

                Console.WriteLine("Table style: " + wordTable.Style);
                Console.WriteLine("Table MarginDefaultTopWidth: " + wordTable.StyleDetails.MarginDefaultTopWidth);

                wordTable.Style = WordTableStyle.GridTable6ColorfulAccent1;

                Console.WriteLine("Table style: " + wordTable.Style);
                Console.WriteLine("Table MarginDefaultTopWidth: " + wordTable.StyleDetails.MarginDefaultTopWidth);

                document.Save(openWord);
            }
        }
    }
}
