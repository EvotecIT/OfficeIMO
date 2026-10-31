using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_BasicTables10_StylesModificationWithCentimeters(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tables demonstrating both twips and centimeters");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables10_StyleModificationWithCentimeters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var title = document.AddParagraph("Table Styles with Twips and Centimeters");
                title.ParagraphAlignment = WordParagraphAlignment.Center;
                title.Bold = true;
                title.FontSize = global::OfficeIMO.OfficeLength.Points(16);
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
                styleDetails1.MarginDefaultTop = global::OfficeIMO.OfficeLength.Points(5.5);
                styleDetails1.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Points(5.5);
                styleDetails1.MarginDefaultLeft = global::OfficeIMO.OfficeLength.Points(5.5);
                styleDetails1.MarginDefaultRight = global::OfficeIMO.OfficeLength.Points(5.5);
                styleDetails1.CellSpacing = global::OfficeIMO.OfficeLength.Points(2.5);

                Console.WriteLine("\nFirst table (Twips):");
                Console.WriteLine("Table MarginDefaultTop: " + styleDetails1.MarginDefaultTop);
                Console.WriteLine("Table top margin (cm): " + styleDetails1.MarginDefaultTop?.ToCentimeters());

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
                styleDetails2.MarginDefaultTop = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                styleDetails2.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                styleDetails2.MarginDefaultLeft = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                styleDetails2.MarginDefaultRight = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                styleDetails2.CellSpacing = global::OfficeIMO.OfficeLength.Centimeters(0.1);

                Console.WriteLine("\nSecond table (Centimeters):");
                Console.WriteLine("Table MarginDefaultTop: " + styleDetails2.MarginDefaultTop);
                Console.WriteLine("Table top margin (cm): " + styleDetails2.MarginDefaultTop?.ToCentimeters());

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
                styleDetails3.MarginDefaultTop = global::OfficeIMO.OfficeLength.Points(6);  // Using points
                styleDetails3.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Centimeters(0.3);  // Using centimeters
                styleDetails3.MarginDefaultLeft = global::OfficeIMO.OfficeLength.Points(7.5);  // Using points
                styleDetails3.MarginDefaultRight = global::OfficeIMO.OfficeLength.Centimeters(0.25);  // Using centimeters

                Console.WriteLine("\nThird table (Mixed):");
                Console.WriteLine("Table MarginDefaultTop: " + styleDetails3.MarginDefaultTop);
                Console.WriteLine("Table MarginDefaultBottom: " + styleDetails3.MarginDefaultBottom);
                Console.WriteLine("Table MarginDefaultLeft: " + styleDetails3.MarginDefaultLeft);
                Console.WriteLine("Table MarginDefaultRight: " + styleDetails3.MarginDefaultRight);

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
                    WordBorderStyle.Double, global::OfficeIMO.OfficeLength.Points(3), Color.Red,  // Outside borders
                    WordBorderStyle.Single, global::OfficeIMO.OfficeLength.Points(1.5), Color.Blue  // Inside borders
                );

                // Set cell spacing using centimeters
                styleDetails4.CellSpacing = global::OfficeIMO.OfficeLength.Centimeters(0.15);

                Console.WriteLine("\nFourth table (Borders):");
                Console.WriteLine("Table CellSpacing: " + styleDetails4.CellSpacing);
                Console.WriteLine("Table cell spacing (cm): " + styleDetails4.CellSpacing?.ToCentimeters());

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