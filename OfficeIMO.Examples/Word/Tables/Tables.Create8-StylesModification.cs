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

                // Set margins for all sides
                wordTable.StyleDetails.MarginDefaultTopWidth = 110;
                wordTable.StyleDetails.MarginDefaultBottomWidth = 110;
                wordTable.StyleDetails.MarginDefaultLeftWidth = 110;
                wordTable.StyleDetails.MarginDefaultRightWidth = 110;
                wordTable.StyleDetails.CellSpacing = 50;

                Console.WriteLine("Table style: " + wordTable.Style);
                Console.WriteLine("Table MarginDefaultTopWidth: " + wordTable.StyleDetails.MarginDefaultTopWidth);
                Console.WriteLine("Table MarginDefaultBottomWidth: " + wordTable.StyleDetails.MarginDefaultBottomWidth);
                Console.WriteLine("Table MarginDefaultLeftWidth: " + wordTable.StyleDetails.MarginDefaultLeftWidth);
                Console.WriteLine("Table MarginDefaultRightWidth: " + wordTable.StyleDetails.MarginDefaultRightWidth);
                Console.WriteLine("Table CellSpacing: " + wordTable.StyleDetails.CellSpacing);

                document.AddParagraph();

                // Create another table with different style and margins
                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Style 2 Test 1";
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Style 2 Test 2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Style 2 Test 3";

                // Set different margins for each side
                wordTable2.StyleDetails.MarginDefaultTopWidth = 120;
                wordTable2.StyleDetails.MarginDefaultBottomWidth = 180;
                wordTable2.StyleDetails.MarginDefaultLeftWidth = 150;
                wordTable2.StyleDetails.MarginDefaultRightWidth = 150;

                // Add custom borders
                TableBorders borders = new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 24 },
                    new BottomBorder() { Val = BorderValues.Double, Size = 24 },
                    new LeftBorder() { Val = BorderValues.Single, Size = 24 },
                    new RightBorder() { Val = BorderValues.Single, Size = 24 },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 12 },
                    new InsideVerticalBorder() { Val = BorderValues.Single, Size = 12 }
                );
                wordTable2.StyleDetails.TableBorders = borders;

                wordTable2.Rows[0].Cells[2].Borders.TopColor = Color.Red;
                wordTable2.Rows[0].Cells[2].Borders.BottomColor = Color.Green;
                wordTable2.Rows[0].Cells[2].Borders.TopSize = 24;
                wordTable2.Rows[0].Cells[2].Borders.TopStyle = BorderValues.Single;

                Console.WriteLine("\nSecond table settings:");
                Console.WriteLine("Table style: " + wordTable2.Style);
                Console.WriteLine("Table MarginDefaultTopWidth: " + wordTable2.StyleDetails.MarginDefaultTopWidth);
                Console.WriteLine("Table MarginDefaultBottomWidth: " + wordTable2.StyleDetails.MarginDefaultBottomWidth);
                Console.WriteLine("Table MarginDefaultLeftWidth: " + wordTable2.StyleDetails.MarginDefaultLeftWidth);
                Console.WriteLine("Table MarginDefaultRightWidth: " + wordTable2.StyleDetails.MarginDefaultRightWidth);

                document.Save(openWord);
            }
        }
    }
}
