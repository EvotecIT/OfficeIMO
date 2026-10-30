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
        internal static void Example_UnifiedTableBorders(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document demonstrating unified table borders functionality");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables9_UnifiedBorders.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var title = document.AddParagraph("Unified Table Border Examples");
                title.ParagraphAlignment = WordParagraphAlignment.Center;
                title.Bold = true;
                title.FontSize = global::OfficeIMO.OfficeLength.Points(16);
                document.AddParagraph();

                // Example 1: Simple uniform borders
                var heading1 = document.AddParagraph("Example 1: Uniform Borders for All Sides");
                heading1.Bold = true;
                document.AddParagraph("This table has the same border style, size, and color applied to all sides.");

                WordTable table1 = document.AddTable(3, 3);
                SetCellText(table1, 0, 0, "Uniform");
                SetCellText(table1, 0, 1, "Border");
                SetCellText(table1, 0, 2, "Example");
                SetCellText(table1, 1, 0, "All");
                SetCellText(table1, 1, 1, "Sides");
                SetCellText(table1, 1, 2, "Same");
                SetCellText(table1, 2, 0, "Single");
                SetCellText(table1, 2, 1, "Blue");
                SetCellText(table1, 2, 2, "1.5 pt");

                // Use the helper method to set uniform borders
                var table1Style = Guard.NotNull(table1.StyleDetails, "Table style details should be available.");
                table1Style.SetBordersForAllSides(
                    WordBorderStyle.Single,
                    global::OfficeIMO.OfficeLength.Points(1.5),
                    Color.Blue
                );

                document.AddParagraph();
                document.AddParagraph();

                // Example 2: Different outside and inside borders
                var heading2 = document.AddParagraph("Example 2: Different Outside/Inside Borders");
                heading2.Bold = true;
                document.AddParagraph("This table has different border styles for outside edges vs. inside grid lines.");

                WordTable table2 = document.AddTable(3, 3);
                SetCellText(table2, 0, 0, "Outside");
                SetCellText(table2, 0, 1, "Different");
                SetCellText(table2, 0, 2, "From Inside");
                SetCellText(table2, 1, 0, "Double");
                SetCellText(table2, 1, 1, "Red");
                SetCellText(table2, 1, 2, "Outside");
                SetCellText(table2, 2, 0, "Single");
                SetCellText(table2, 2, 1, "Blue");
                SetCellText(table2, 2, 2, "Inside");

                // Using the helper method to set different border styles for outside vs inside
                var table2Style = Guard.NotNull(table2.StyleDetails, "Table style details should be available.");
                table2Style.SetBordersOutsideInside(
                    WordBorderStyle.Double, global::OfficeIMO.OfficeLength.Points(2), Color.Red,  // Outside borders
                    WordBorderStyle.Single, global::OfficeIMO.OfficeLength.Points(1), Color.Blue   // Inside borders
                );

                document.AddParagraph();
                document.AddParagraph();

                // Example 3: Custom borders for each side
                var heading3 = document.AddParagraph("Example 3: Custom Border for Each Side");
                heading3.Bold = true;
                document.AddParagraph("This table has different border styles for each side.");

                WordTable table3 = document.AddTable(3, 3);
                SetCellText(table3, 0, 0, "Top");
                SetCellText(table3, 0, 1, "DotDash");
                SetCellText(table3, 0, 2, "Green");
                SetCellText(table3, 1, 0, "Left");
                SetCellText(table3, 1, 1, "Custom");
                SetCellText(table3, 1, 2, "Right");
                SetCellText(table3, 2, 0, "Triple");
                SetCellText(table3, 2, 1, "Dotted");
                SetCellText(table3, 2, 2, "Bottom");

                // Using the custom border method to set different styles for each side
                var table3Style = Guard.NotNull(table3.StyleDetails, "Table style details should be available.");
                table3Style.SetCustomBorders(
                    topStyle: WordBorderStyle.DotDash, topSize: global::OfficeIMO.OfficeLength.Points(2), topColor: Color.Green,
                    bottomStyle: WordBorderStyle.Triple, bottomSize: global::OfficeIMO.OfficeLength.Points(2), bottomColor: Color.Purple,
                    leftStyle: WordBorderStyle.Thick, leftSize: global::OfficeIMO.OfficeLength.Points(1.5), leftColor: Color.Orange,
                    rightStyle: WordBorderStyle.Wave, rightSize: global::OfficeIMO.OfficeLength.Points(1.5), rightColor: Color.Red,
                    insideHStyle: WordBorderStyle.Dotted, insideHSize: global::OfficeIMO.OfficeLength.Points(1), insideHColor: Color.Gray,
                    insideVStyle: WordBorderStyle.Dotted, insideVSize: global::OfficeIMO.OfficeLength.Points(1), insideVColor: Color.Gray
                );

                document.AddParagraph();
                document.AddParagraph();

                // Example 4: Apply table borders to individual cells
                var heading4 = document.AddParagraph("Example 4: Apply Table Borders to All Cells");
                heading4.Bold = true;
                document.AddParagraph("This table has borders defined at the table level and then applied to all individual cells.");

                WordTable table4 = document.AddTable(3, 3);
                SetCellText(table4, 0, 0, "Table");
                SetCellText(table4, 0, 1, "Borders");
                SetCellText(table4, 0, 2, "Applied");
                SetCellText(table4, 1, 0, "To");
                SetCellText(table4, 1, 1, "All");
                SetCellText(table4, 1, 2, "Cells");
                SetCellText(table4, 2, 0, "Consistent");
                SetCellText(table4, 2, 1, "Border");
                SetCellText(table4, 2, 2, "Styling");

                // Set table borders first
                var table4Style = Guard.NotNull(table4.StyleDetails, "Table style details should be available.");
                table4Style.SetBordersOutsideInside(
                    WordBorderStyle.Double, global::OfficeIMO.OfficeLength.Points(3), Color.DarkBlue,   // Outside borders
                    WordBorderStyle.Single, global::OfficeIMO.OfficeLength.Points(1.5), Color.DarkGreen   // Inside borders
                );

                // Then apply those borders to all cells
                table4Style.ApplyBordersToAllCells();

                document.AddParagraph();
                document.AddParagraph();

                // Example 5: Get border properties
                var heading5 = document.AddParagraph("Example 5: Get Border Properties");
                heading5.Bold = true;
                document.AddParagraph("This table demonstrates retrieving border properties using the GetBorderProperties method.");

                WordTable table5 = document.AddTable(3, 3);

                // Set some border styles first
                var table5Style = Guard.NotNull(table5.StyleDetails, "Table style details should be available.");
                table5Style.SetCustomBorders(
                    topStyle: WordBorderStyle.Thick, topSize: global::OfficeIMO.OfficeLength.Points(2), topColor: Color.DarkRed,
                    bottomStyle: WordBorderStyle.Thick, bottomSize: global::OfficeIMO.OfficeLength.Points(2), bottomColor: Color.DarkRed,
                    leftStyle: WordBorderStyle.Single, leftSize: global::OfficeIMO.OfficeLength.Points(1), leftColor: Color.DarkBlue,
                    rightStyle: WordBorderStyle.Single, rightSize: global::OfficeIMO.OfficeLength.Points(1), rightColor: Color.DarkBlue
                );

                // Retrieve the properties and display them
                var topBorderProps = table5Style.GetBorderProperties(WordTableBorderSide.Top);
                var bottomBorderProps = table5Style.GetBorderProperties(WordTableBorderSide.Bottom);
                var leftBorderProps = table5Style.GetBorderProperties(WordTableBorderSide.Left);
                var rightBorderProps = table5Style.GetBorderProperties(WordTableBorderSide.Right);

                SetCellText(table5, 0, 0, "Top");
                SetCellText(table5, 0, 1, $"Style: {topBorderProps.Style}");
                SetCellText(table5, 0, 2, $"Size: {topBorderProps.Size}");

                SetCellText(table5, 1, 0, "Bottom");
                SetCellText(table5, 1, 1, $"Style: {bottomBorderProps.Style}");
                SetCellText(table5, 1, 2, $"Size: {bottomBorderProps.Size}");

                SetCellText(table5, 2, 0, "Left/Right");
                SetCellText(table5, 2, 1, $"Style: {leftBorderProps.Style}");
                SetCellText(table5, 2, 2, $"Size: {rightBorderProps.Size}");

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
