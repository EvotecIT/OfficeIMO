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
using Color = SixLabors.ImageSharp.Color;

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
                SetCellText(wordTable, 0, 0, "Test 1");
                SetCellText(wordTable, 1, 0, "Test 2");
                SetCellText(wordTable, 2, 0, "Test 3");

                // Set margins for all sides
                var styleDetails1 = Guard.NotNull(wordTable.StyleDetails, "Table style details should be available.");
                styleDetails1.MarginDefaultTopWidth = 110;
                styleDetails1.MarginDefaultBottomWidth = 110;
                styleDetails1.MarginDefaultLeftWidth = 110;
                styleDetails1.MarginDefaultRightWidth = 110;
                styleDetails1.CellSpacing = 50;

                Console.WriteLine("Table style: " + wordTable.Style);
                Console.WriteLine("Table MarginDefaultTopWidth: " + styleDetails1.MarginDefaultTopWidth);
                Console.WriteLine("Table MarginDefaultBottomWidth: " + styleDetails1.MarginDefaultBottomWidth);
                Console.WriteLine("Table MarginDefaultLeftWidth: " + styleDetails1.MarginDefaultLeftWidth);
                Console.WriteLine("Table MarginDefaultRightWidth: " + styleDetails1.MarginDefaultRightWidth);
                Console.WriteLine("Table CellSpacing: " + styleDetails1.CellSpacing);

                document.AddParagraph();

                // Create another table with different style and margins
                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                SetCellText(wordTable2, 0, 0, "Style 2 Test 1");
                SetCellText(wordTable2, 1, 0, "Style 2 Test 2");
                SetCellText(wordTable2, 2, 0, "Style 2 Test 3");

                // Set different margins for each side
                var styleDetails2 = Guard.NotNull(wordTable2.StyleDetails, "Table style details should be available.");
                styleDetails2.MarginDefaultTopWidth = 120;
                styleDetails2.MarginDefaultBottomWidth = 180;
                styleDetails2.MarginDefaultLeftWidth = 150;
                styleDetails2.MarginDefaultRightWidth = 150;

                // Add custom borders
                TableBorders borders = new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 24 },
                    new BottomBorder() { Val = BorderValues.Double, Size = 24 },
                    new LeftBorder() { Val = BorderValues.Single, Size = 24 },
                    new RightBorder() { Val = BorderValues.Single, Size = 24 },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 12 },
                    new InsideVerticalBorder() { Val = BorderValues.Single, Size = 12 }
                );
                styleDetails2.TableBorders = borders;

                wordTable2.Rows[0].Cells[2].Borders.TopColor = Color.Red;
                wordTable2.Rows[0].Cells[2].Borders.BottomColor = Color.Green;
                wordTable2.Rows[0].Cells[2].Borders.TopSize = 24;
                wordTable2.Rows[0].Cells[2].Borders.TopStyle = BorderValues.Single;

                Console.WriteLine("\nSecond table settings:");
                Console.WriteLine("Table style: " + wordTable2.Style);
                Console.WriteLine("Table MarginDefaultTopWidth: " + styleDetails2.MarginDefaultTopWidth);
                Console.WriteLine("Table MarginDefaultBottomWidth: " + styleDetails2.MarginDefaultBottomWidth);
                Console.WriteLine("Table MarginDefaultLeftWidth: " + styleDetails2.MarginDefaultLeftWidth);
                Console.WriteLine("Table MarginDefaultRightWidth: " + styleDetails2.MarginDefaultRightWidth);

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
