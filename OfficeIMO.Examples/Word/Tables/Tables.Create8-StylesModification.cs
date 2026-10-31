using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_BasicTables8_StylesModification(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tables");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables8_StyleModification.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
                document.AddParagraph();

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                SetCellText(wordTable, 0, 0, "Test 1");
                SetCellText(wordTable, 1, 0, "Test 2");
                SetCellText(wordTable, 2, 0, "Test 3");

                // Set margins for all sides
                var styleDetails1 = Guard.NotNull(wordTable.StyleDetails, "Table style details should be available.");
                styleDetails1.MarginDefaultTop = global::OfficeIMO.OfficeLength.Points(5.5);
                styleDetails1.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Points(5.5);
                styleDetails1.MarginDefaultLeft = global::OfficeIMO.OfficeLength.Points(5.5);
                styleDetails1.MarginDefaultRight = global::OfficeIMO.OfficeLength.Points(5.5);
                styleDetails1.CellSpacing = global::OfficeIMO.OfficeLength.Points(2.5);

                Console.WriteLine("Table style: " + wordTable.Style);
                Console.WriteLine("Table MarginDefaultTop: " + styleDetails1.MarginDefaultTop);
                Console.WriteLine("Table MarginDefaultBottom: " + styleDetails1.MarginDefaultBottom);
                Console.WriteLine("Table MarginDefaultLeft: " + styleDetails1.MarginDefaultLeft);
                Console.WriteLine("Table MarginDefaultRight: " + styleDetails1.MarginDefaultRight);
                Console.WriteLine("Table CellSpacing: " + styleDetails1.CellSpacing);

                document.AddParagraph();

                // Create another table with different style and margins
                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                SetCellText(wordTable2, 0, 0, "Style 2 Test 1");
                SetCellText(wordTable2, 1, 0, "Style 2 Test 2");
                SetCellText(wordTable2, 2, 0, "Style 2 Test 3");

                // Set different margins for each side
                var styleDetails2 = Guard.NotNull(wordTable2.StyleDetails, "Table style details should be available.");
                styleDetails2.MarginDefaultTop = global::OfficeIMO.OfficeLength.Points(6);
                styleDetails2.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Points(9);
                styleDetails2.MarginDefaultLeft = global::OfficeIMO.OfficeLength.Points(7.5);
                styleDetails2.MarginDefaultRight = global::OfficeIMO.OfficeLength.Points(7.5);

                // Add custom borders
                styleDetails2.SetCustomBorders(
                    topStyle: WordBorderStyle.Single, topSize: global::OfficeIMO.OfficeLength.Points(3),
                    bottomStyle: WordBorderStyle.Double, bottomSize: global::OfficeIMO.OfficeLength.Points(3),
                    leftStyle: WordBorderStyle.Single, leftSize: global::OfficeIMO.OfficeLength.Points(3),
                    rightStyle: WordBorderStyle.Single, rightSize: global::OfficeIMO.OfficeLength.Points(3),
                    insideHStyle: WordBorderStyle.Single, insideHSize: global::OfficeIMO.OfficeLength.Points(1.5),
                    insideVStyle: WordBorderStyle.Single, insideVSize: global::OfficeIMO.OfficeLength.Points(1.5));

                wordTable2.Rows[0].Cells[2].Borders.TopColor = Color.Red;
                wordTable2.Rows[0].Cells[2].Borders.BottomColor = Color.Green;
                wordTable2.Rows[0].Cells[2].Borders.TopSize = OfficeLength.Points(3);
                wordTable2.Rows[0].Cells[2].Borders.TopStyle = WordBorderStyle.Single;

                Console.WriteLine("\nSecond table settings:");
                Console.WriteLine("Table style: " + wordTable2.Style);
                Console.WriteLine("Table MarginDefaultTop: " + styleDetails2.MarginDefaultTop);
                Console.WriteLine("Table MarginDefaultBottom: " + styleDetails2.MarginDefaultBottom);
                Console.WriteLine("Table MarginDefaultLeft: " + styleDetails2.MarginDefaultLeft);
                Console.WriteLine("Table MarginDefaultRight: " + styleDetails2.MarginDefaultRight);

                document.Save();
                if (openWord) document.OpenInApplication();

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
