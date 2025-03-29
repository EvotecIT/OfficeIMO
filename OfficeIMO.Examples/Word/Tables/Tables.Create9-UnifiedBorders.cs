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
        internal static void Example_UnifiedTableBorders(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document demonstrating unified table borders functionality");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables9_UnifiedBorders.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var title = document.AddParagraph("Unified Table Border Examples");
                title.ParagraphAlignment = JustificationValues.Center;
                title.Bold = true;
                title.FontSize = 16;
                document.AddParagraph();

                // Example 1: Simple uniform borders
                var heading1 = document.AddParagraph("Example 1: Uniform Borders for All Sides");
                heading1.Bold = true;
                document.AddParagraph("This table has the same border style, size, and color applied to all sides.");

                WordTable table1 = document.AddTable(3, 3);
                table1.Rows[0].Cells[0].Paragraphs[0].Text = "Uniform";
                table1.Rows[0].Cells[1].Paragraphs[0].Text = "Border";
                table1.Rows[0].Cells[2].Paragraphs[0].Text = "Example";
                table1.Rows[1].Cells[0].Paragraphs[0].Text = "All";
                table1.Rows[1].Cells[1].Paragraphs[0].Text = "Sides";
                table1.Rows[1].Cells[2].Paragraphs[0].Text = "Same";
                table1.Rows[2].Cells[0].Paragraphs[0].Text = "Single";
                table1.Rows[2].Cells[1].Paragraphs[0].Text = "Blue";
                table1.Rows[2].Cells[2].Paragraphs[0].Text = "Size 12";

                // Using the new helper method to set uniform borders
                table1.StyleDetails.SetBordersForAllSides(
                    BorderValues.Single,
                    12U,
                    Color.Blue
                );

                document.AddParagraph();
                document.AddParagraph();

                // Example 2: Different outside and inside borders
                var heading2 = document.AddParagraph("Example 2: Different Outside/Inside Borders");
                heading2.Bold = true;
                document.AddParagraph("This table has different border styles for outside edges vs. inside grid lines.");

                WordTable table2 = document.AddTable(3, 3);
                table2.Rows[0].Cells[0].Paragraphs[0].Text = "Outside";
                table2.Rows[0].Cells[1].Paragraphs[0].Text = "Different";
                table2.Rows[0].Cells[2].Paragraphs[0].Text = "From Inside";
                table2.Rows[1].Cells[0].Paragraphs[0].Text = "Double";
                table2.Rows[1].Cells[1].Paragraphs[0].Text = "Red";
                table2.Rows[1].Cells[2].Paragraphs[0].Text = "Outside";
                table2.Rows[2].Cells[0].Paragraphs[0].Text = "Single";
                table2.Rows[2].Cells[1].Paragraphs[0].Text = "Blue";
                table2.Rows[2].Cells[2].Paragraphs[0].Text = "Inside";

                // Using the helper method to set different border styles for outside vs inside
                table2.StyleDetails.SetBordersOutsideInside(
                    BorderValues.Double, 16U, Color.Red,  // Outside borders
                    BorderValues.Single, 8U, Color.Blue   // Inside borders
                );

                document.AddParagraph();
                document.AddParagraph();

                // Example 3: Custom borders for each side
                var heading3 = document.AddParagraph("Example 3: Custom Border for Each Side");
                heading3.Bold = true;
                document.AddParagraph("This table has different border styles for each side.");

                WordTable table3 = document.AddTable(3, 3);
                table3.Rows[0].Cells[0].Paragraphs[0].Text = "Top";
                table3.Rows[0].Cells[1].Paragraphs[0].Text = "DotDash";
                table3.Rows[0].Cells[2].Paragraphs[0].Text = "Green";
                table3.Rows[1].Cells[0].Paragraphs[0].Text = "Left";
                table3.Rows[1].Cells[1].Paragraphs[0].Text = "Custom";
                table3.Rows[1].Cells[2].Paragraphs[0].Text = "Right";
                table3.Rows[2].Cells[0].Paragraphs[0].Text = "Triple";
                table3.Rows[2].Cells[1].Paragraphs[0].Text = "Dotted";
                table3.Rows[2].Cells[2].Paragraphs[0].Text = "Bottom";

                // Using the custom border method to set different styles for each side
                table3.StyleDetails.SetCustomBorders(
                    topStyle: BorderValues.DotDash, topSize: 16U, topColor: Color.Green,
                    bottomStyle: BorderValues.Triple, bottomSize: 16U, bottomColor: Color.Purple,
                    leftStyle: BorderValues.Thick, leftSize: 12U, leftColor: Color.Orange,
                    rightStyle: BorderValues.Wave, rightSize: 12U, rightColor: Color.Red,
                    insideHStyle: BorderValues.Dotted, insideHSize: 8U, insideHColor: Color.Gray,
                    insideVStyle: BorderValues.Dotted, insideVSize: 8U, insideVColor: Color.Gray
                );

                document.AddParagraph();
                document.AddParagraph();

                // Example 4: Apply table borders to individual cells
                var heading4 = document.AddParagraph("Example 4: Apply Table Borders to All Cells");
                heading4.Bold = true;
                document.AddParagraph("This table has borders defined at the table level and then applied to all individual cells.");

                WordTable table4 = document.AddTable(3, 3);
                table4.Rows[0].Cells[0].Paragraphs[0].Text = "Table";
                table4.Rows[0].Cells[1].Paragraphs[0].Text = "Borders";
                table4.Rows[0].Cells[2].Paragraphs[0].Text = "Applied";
                table4.Rows[1].Cells[0].Paragraphs[0].Text = "To";
                table4.Rows[1].Cells[1].Paragraphs[0].Text = "All";
                table4.Rows[1].Cells[2].Paragraphs[0].Text = "Cells";
                table4.Rows[2].Cells[0].Paragraphs[0].Text = "Consistent";
                table4.Rows[2].Cells[1].Paragraphs[0].Text = "Border";
                table4.Rows[2].Cells[2].Paragraphs[0].Text = "Styling";

                // Set table borders first
                table4.StyleDetails.SetBordersOutsideInside(
                    BorderValues.Double, 24U, Color.DarkBlue,   // Outside borders
                    BorderValues.Single, 12U, Color.DarkGreen   // Inside borders
                );

                // Then apply those borders to all cells
                table4.StyleDetails.ApplyBordersToAllCells();

                document.AddParagraph();
                document.AddParagraph();

                // Example 5: Get border properties
                var heading5 = document.AddParagraph("Example 5: Get Border Properties");
                heading5.Bold = true;
                document.AddParagraph("This table demonstrates retrieving border properties using the GetBorderProperties method.");

                WordTable table5 = document.AddTable(3, 3);

                // Set some border styles first
                table5.StyleDetails.SetCustomBorders(
                    topStyle: BorderValues.Thick, topSize: 16U, topColor: Color.DarkRed,
                    bottomStyle: BorderValues.Thick, bottomSize: 16U, bottomColor: Color.DarkRed,
                    leftStyle: BorderValues.Single, leftSize: 8U, leftColor: Color.DarkBlue,
                    rightStyle: BorderValues.Single, rightSize: 8U, rightColor: Color.DarkBlue
                );

                // Retrieve the properties and display them
                var topBorderProps = table5.StyleDetails.GetBorderProperties(BorderSide.Top);
                var bottomBorderProps = table5.StyleDetails.GetBorderProperties(BorderSide.Bottom);
                var leftBorderProps = table5.StyleDetails.GetBorderProperties(BorderSide.Left);
                var rightBorderProps = table5.StyleDetails.GetBorderProperties(BorderSide.Right);

                table5.Rows[0].Cells[0].Paragraphs[0].Text = "Top";
                table5.Rows[0].Cells[1].Paragraphs[0].Text = $"Style: {topBorderProps.Style}";
                table5.Rows[0].Cells[2].Paragraphs[0].Text = $"Size: {topBorderProps.Size}";

                table5.Rows[1].Cells[0].Paragraphs[0].Text = "Bottom";
                table5.Rows[1].Cells[1].Paragraphs[0].Text = $"Style: {bottomBorderProps.Style}";
                table5.Rows[1].Cells[2].Paragraphs[0].Text = $"Size: {bottomBorderProps.Size}";

                table5.Rows[2].Cells[0].Paragraphs[0].Text = "Left/Right";
                table5.Rows[2].Cells[1].Paragraphs[0].Text = $"Style: {leftBorderProps.Style}";
                table5.Rows[2].Cells[2].Paragraphs[0].Text = $"Size: {rightBorderProps.Size}";

                document.Save(openWord);
            }
        }
    }
}