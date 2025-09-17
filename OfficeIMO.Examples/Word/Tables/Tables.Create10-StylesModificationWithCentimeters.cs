using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_BasicTables10_StylesModificationWithCentimeters(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tables demonstrating both twips and centimeters");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables10_StyleModificationWithCentimeters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var title = document.AddParagraph("Table Styles with Twips and Centimeters");
                title.ParagraphAlignment = JustificationValues.Center;
                title.Bold = true;
                title.FontSize = 16;
                document.AddParagraph();

                // Example 1: Using Twips (Original Method)
                var heading1 = document.AddParagraph("Example 1: Setting Margins Using Twips");
                heading1.Bold = true;
                document.AddParagraph("This table uses the original twips-based method for setting margins.");

                WordTable wordTable1 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                SetCellText(wordTable1, 0, 0, "Twips Example");
                SetCellText(wordTable1, 1, 0, "Using Original Method");
                SetCellText(wordTable1, 2, 0, "All sides 110 twips");

                // Set margins using twips
                var styleDetails1 = Guard.NotNull(wordTable1.StyleDetails, "Table style details should be available.");
                styleDetails1.MarginDefaultTopWidth = 110;
                styleDetails1.MarginDefaultBottomWidth = 110;
                styleDetails1.MarginDefaultLeftWidth = 110;
                styleDetails1.MarginDefaultRightWidth = 110;
                styleDetails1.CellSpacing = 50;

                Console.WriteLine("\nFirst table (Twips):");
                Console.WriteLine("Table MarginDefaultTopWidth: " + styleDetails1.MarginDefaultTopWidth);
                Console.WriteLine("Table MarginDefaultTopCentimeters: " + styleDetails1.MarginDefaultTopCentimeters);

                document.AddParagraph();

                // Example 2: Using Centimeters (New Method)
                var heading2 = document.AddParagraph("Example 2: Setting Margins Using Centimeters");
                heading2.Bold = true;
                document.AddParagraph("This table uses the new centimeters-based method for setting margins.");

                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                SetCellText(wordTable2, 0, 0, "Centimeters Example");
                SetCellText(wordTable2, 1, 0, "Using New Method");
                SetCellText(wordTable2, 2, 0, "All sides 0.2 cm");

                // Set margins using centimeters
                var styleDetails2 = Guard.NotNull(wordTable2.StyleDetails, "Table style details should be available.");
                styleDetails2.MarginDefaultTopCentimeters = 0.2;
                styleDetails2.MarginDefaultBottomCentimeters = 0.2;
                styleDetails2.MarginDefaultLeftCentimeters = 0.2;
                styleDetails2.MarginDefaultRightCentimeters = 0.2;
                styleDetails2.CellSpacingCentimeters = 0.1;

                Console.WriteLine("\nSecond table (Centimeters):");
                Console.WriteLine("Table MarginDefaultTopCentimeters: " + styleDetails2.MarginDefaultTopCentimeters);
                Console.WriteLine("Table MarginDefaultTopWidth: " + styleDetails2.MarginDefaultTopWidth);

                document.AddParagraph();

                // Example 3: Mixed Approach with Different Sides
                var heading3 = document.AddParagraph("Example 3: Mixed Approach - Different Units for Different Sides");
                heading3.Bold = true;
                document.AddParagraph("This table demonstrates using both twips and centimeters for different sides.");

                WordTable wordTable3 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                SetCellText(wordTable3, 0, 0, "Mixed Units Example");
                SetCellText(wordTable3, 1, 0, "Using Both Methods");
                SetCellText(wordTable3, 2, 0, "Different sides");

                // Mix of twips and centimeters
                var styleDetails3 = Guard.NotNull(wordTable3.StyleDetails, "Table style details should be available.");
                styleDetails3.MarginDefaultTopWidth = 120;  // Using twips
                styleDetails3.MarginDefaultBottomCentimeters = 0.3;  // Using centimeters
                styleDetails3.MarginDefaultLeftWidth = 150;  // Using twips
                styleDetails3.MarginDefaultRightCentimeters = 0.25;  // Using centimeters

                Console.WriteLine("\nThird table (Mixed):");
                Console.WriteLine("Table MarginDefaultTopWidth: " + styleDetails3.MarginDefaultTopWidth);
                Console.WriteLine("Table MarginDefaultBottomCentimeters: " + styleDetails3.MarginDefaultBottomCentimeters);
                Console.WriteLine("Table MarginDefaultLeftWidth: " + styleDetails3.MarginDefaultLeftWidth);
                Console.WriteLine("Table MarginDefaultRightCentimeters: " + styleDetails3.MarginDefaultRightCentimeters);

                document.AddParagraph();

                // Example 4: Border Styles with Both Units
                var heading4 = document.AddParagraph("Example 4: Border Styles with Different Units");
                heading4.Bold = true;
                document.AddParagraph("This table demonstrates border styles with different units for spacing.");

                WordTable wordTable4 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                SetCellText(wordTable4, 0, 0, "Border Styles");
                SetCellText(wordTable4, 1, 0, "With Different Units");
                SetCellText(wordTable4, 2, 0, "For Spacing");

                // Set borders with different units
                var styleDetails4 = Guard.NotNull(wordTable4.StyleDetails, "Table style details should be available.");
                styleDetails4.SetBordersOutsideInside(
                    BorderValues.Double, 24U, Color.Red,  // Outside borders
                    BorderValues.Single, 12U, Color.Blue  // Inside borders
                );

                // Set cell spacing using centimeters
                styleDetails4.CellSpacingCentimeters = 0.15;

                Console.WriteLine("\nFourth table (Borders):");
                Console.WriteLine("Table CellSpacingCentimeters: " + styleDetails4.CellSpacingCentimeters);
                Console.WriteLine("Table CellSpacing: " + styleDetails4.CellSpacing);

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