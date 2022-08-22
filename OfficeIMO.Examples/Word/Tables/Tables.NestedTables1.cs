using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_NestedTables(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with nested tables");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Nested Tables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Lets add table ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable.Rows[2].Cells[0].AddTable(2, 2, WordTableStyle.GridTable2Accent2);


                document.Save(openWord);
            }
        }
    }
}
