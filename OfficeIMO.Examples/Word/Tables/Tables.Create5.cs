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
        internal static void Example_TablesWidthAndAlignment(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with width and alignment");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Table Alignment.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Lets add table with some alignment ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable.FirstRow.FirstCell.Paragraphs[0].AddComment("Adam Kłys", "AK", "Test comment for paragraph within a Table");

                wordTable.FirstRow.FirstCell.ShadingFillColor = Color.Blue;
                wordTable.Rows[1].FirstCell.ShadingFillColor = Color.Red;

                wordTable.LastRow.FirstCell.ShadingPattern = ShadingPatternValues.Percent20;


                wordTable.AddComment("Przemysław Kłys", "PK", "This is a table, and we just added comment to a whole table");

                wordTable.LastRow.LastCell.Paragraphs[0].Text = "Last Cell";

                wordTable.WidthType = TableWidthUnitValues.Pct;
                wordTable.Width = "3000";
                wordTable.Alignment = TableRowAlignmentValues.Center;

                document.Save(openWord);
            }
        }
    }
}
