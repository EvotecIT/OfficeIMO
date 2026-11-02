using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_TablesWidthAndAlignment(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with width and alignment");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Table Alignment.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Lets add table with some alignment ");
                paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
                paragraph.Bold = true;
                paragraph.Underline = WordUnderlineStyle.DotDash;

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable.FirstRow.FirstCell.Paragraphs[0].AddComment("Adam Kłys", "AK", "Test comment for paragraph within a Table");

                wordTable.FirstRow.FirstCell.ShadingFillColor = Color.Blue;
                wordTable.Rows[1].FirstCell.ShadingFillColor = Color.Red;

                wordTable.LastRow.FirstCell.ShadingPattern = WordShadingPattern.Percent20;


                wordTable.AddComment("Przemysław Kłys", "PK", "This is a table, and we just added comment to a whole table");

                wordTable.LastRow.LastCell.Paragraphs[0].Text = "Last Cell";

                wordTable.Width = WordTableWidth.Percent(100);
                wordTable.Alignment = WordTableAlignment.Center;

                wordTable.Title = "This is title";
                wordTable.Description = "Description of table";


                var paragraph1 = document.AddParagraph("Lets add another table showing text wrapping around");

                WordTable wordTable1 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable1.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable1.Width = WordTableWidth.Percent(60);

                wordTable1.AllowTextWrap = true;

                var paragraph2 = document.AddParagraph("This paragraph should continue but next to to the table");

                document.AddParagraph();
                document.AddParagraph();

                var paragraph3 = document.AddParagraph("Lets add another table showing AutoFit");

                WordTable wordTable2 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable2.ColumnWidth = new List<WordTableWidth?> { WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(85.8)), WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(190.85)), WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(15)), WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(150)) };
                wordTable2.RowHeight = new List<OfficeIMO.OfficeLength?> { OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeLength.Points(15), OfficeIMO.OfficeLength.Points(25), OfficeIMO.OfficeLength.Points(10) };

                // add a cell to 3rd row
                WordTableCell cell = wordTable2.Rows[2].AddCell();
                cell.Paragraphs[0].Text = "This cell is outside a bit";
                cell.TextDirection = WordTextDirection.TopToBottomLeftToRightRotated;

                wordTable2.LayoutMode = WordTableLayoutMode.Fixed;

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
