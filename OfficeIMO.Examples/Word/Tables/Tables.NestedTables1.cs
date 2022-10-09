using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

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

                var table1 = wordTable.Rows[0].Cells[0].AddTable(3, 2, WordTableStyle.GridTable2Accent2);

                var table2 = wordTable.Rows[0].Cells[1].AddTable(3, 2, WordTableStyle.GridTable2Accent5, true);

                Console.WriteLine("Table has nested tables: " + document.Tables[0].HasNestedTables);
                Console.WriteLine("Table1 is nested table: " + document.Tables[0].IsNestedTable);

                Console.WriteLine("Table2 is nested table: " + table1.IsNestedTable);
                Console.WriteLine("Table3 is nested table: " + table2.IsNestedTable);


                wordTable.NestedTables[0].Rows[0].Cells[0].Paragraphs[0].Text = "Nested table 1 / 1st row / 1st cell";

                wordTable.NestedTables[1].Rows[0].Cells[1].Paragraphs[0].Text = "Nested table 2 - 1st row / 2nd cell";

                var paragraph1 = document.AddParagraph("Lets add table number 2 ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                WordTable wordTable1 = document.AddTable(5, 5, WordTableStyle.GridTable1LightAccent1);
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 1.2";
                wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "Test 1.3";
                wordTable1.Rows[3].Cells[0].Paragraphs[0].Text = "Test 1.4";
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1.5";


                var paragraph2 = document.AddParagraph("Lets add table number 3 ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                WordTable wordTable2 = document.AddTable(5, 5, WordTableStyle.GridTable1LightAccent1);
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2.2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 2.3";
                wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 2.4";
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 2.5";


                var table3 = wordTable2.Rows[0].Cells[0].AddTable(3, 2, WordTableStyle.GridTable2Accent2);

                var table4 = wordTable2.Rows[0].Cells[1].AddTable(3, 2, WordTableStyle.GridTable2Accent5, true);

                var i = 0;
                foreach (var table in document.TablesIncludingNestedTables) {
                    Console.WriteLine("Table " + i + " out of " + document.TablesIncludingNestedTables.Count);
                    if (table.IsNestedTable) {
                        Console.WriteLine("Nested table Rows count: " + table.RowsCount);
                        Console.WriteLine("Nested table Parent Rows count: " + table.ParentTable.RowsCount);
                    } else {
                        Console.WriteLine("Skipped table, not nested table");
                    }

                    i++;
                }

                document.Save(openWord);
            }
        }
    }
}
