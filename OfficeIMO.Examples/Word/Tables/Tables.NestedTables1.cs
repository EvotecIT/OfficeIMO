using System;
using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
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
                SetCellText(wordTable, 0, 0, "Test 1");
                SetCellText(wordTable, 1, 0, "Test 2");
                SetCellText(wordTable, 2, 0, "Test 3");
                SetCellText(wordTable, 3, 0, "Test 4");

                var table1 = wordTable.Rows[0].Cells[0].AddTable(3, 2, WordTableStyle.GridTable2Accent2);

                var table2 = wordTable.Rows[0].Cells[1].AddTable(3, 2, WordTableStyle.GridTable2Accent5, true);

                if (document.Tables.Count > 0) {
                    Console.WriteLine("Table has nested tables: " + document.Tables[0].HasNestedTables);
                    Console.WriteLine("Table1 is nested table: " + document.Tables[0].IsNestedTable);
                }

                Console.WriteLine("Table2 is nested table: " + table1.IsNestedTable);
                Console.WriteLine("Table3 is nested table: " + table2.IsNestedTable);


                if (wordTable.NestedTables.Count > 0) {
                    var nestedTable1 = Guard.GetRequiredItem(wordTable.NestedTables, 0, "Nested table is missing.");
                    SetCellText(nestedTable1, 0, 0, "Nested table 1 / 1st row / 1st cell");
                }

                if (wordTable.NestedTables.Count > 1) {
                    var nestedTable2 = Guard.GetRequiredItem(wordTable.NestedTables, 1, "Nested table is missing.");
                    SetCellText(nestedTable2, 0, 1, "Nested table 2 - 1st row / 2nd cell");
                }

                var paragraph1 = document.AddParagraph("Lets add table number 2 ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                WordTable wordTable1 = document.AddTable(5, 5, WordTableStyle.GridTable1LightAccent1);
                SetCellText(wordTable1, 1, 0, "Test 1.2");
                SetCellText(wordTable1, 2, 0, "Test 1.3");
                SetCellText(wordTable1, 3, 0, "Test 1.4");
                SetCellText(wordTable1, 0, 0, "Test 1.5");


                var paragraph2 = document.AddParagraph("Lets add table number 3 ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                WordTable wordTable2 = document.AddTable(5, 5, WordTableStyle.GridTable1LightAccent1);
                SetCellText(wordTable2, 1, 0, "Test 2.2");
                SetCellText(wordTable2, 2, 0, "Test 2.3");
                SetCellText(wordTable2, 3, 0, "Test 2.4");
                SetCellText(wordTable2, 0, 0, "Test 2.5");


                var table3 = wordTable2.Rows[0].Cells[0].AddTable(3, 2, WordTableStyle.GridTable2Accent2);

                var table4 = wordTable2.Rows[0].Cells[1].AddTable(3, 2, WordTableStyle.GridTable2Accent5, true);

                var i = 0;
                foreach (var table in document.TablesIncludingNestedTables) {
                    Console.WriteLine("Table " + i + " out of " + document.TablesIncludingNestedTables.Count);
                    if (table.IsNestedTable) {
                        Console.WriteLine("Nested table Rows count: " + table.RowsCount);
                        var parentTable = Guard.NotNull(table.ParentTable, "Nested table should expose its parent table.");
                        Console.WriteLine("Nested table Parent Rows count: " + parentTable.RowsCount);
                    } else {
                        Console.WriteLine("Skipped table, not nested table");
                    }

                    i++;
                }

                document.Save(openWord);

                static void SetCellText(WordTable table, int rowIndex, int columnIndex, string text) {
                    var row = Guard.GetRequiredItem(table.Rows, rowIndex, $"Table must contain row index {rowIndex}.");
                    var cell = Guard.GetRequiredItem(row.Cells, columnIndex, $"Row must contain cell index {columnIndex}.");
                    var paragraph = cell.Paragraphs.FirstOrDefault() ?? cell.AddParagraph();
                    paragraph.Text = text;
                }
            }
        }
    }
}
