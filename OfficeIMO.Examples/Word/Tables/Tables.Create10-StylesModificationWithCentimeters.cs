using DocumentFormat.OpenXml.Wordprocessing;
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
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Twips Example";
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Using Original Method";
                wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "All sides 110 twips";

                // Set margins using twips
                wordTable1.StyleDetails.MarginDefaultTopWidth = 110;
                wordTable1.StyleDetails.MarginDefaultBottomWidth = 110;
                wordTable1.StyleDetails.MarginDefaultLeftWidth = 110;
                wordTable1.StyleDetails.MarginDefaultRightWidth = 110;
                wordTable1.StyleDetails.CellSpacing = 50;

                Console.WriteLine("\nFirst table (Twips):");
                Console.WriteLine("Table MarginDefaultTopWidth: " + wordTable1.StyleDetails.MarginDefaultTopWidth);
                Console.WriteLine("Table MarginDefaultTopCentimeters: " + wordTable1.StyleDetails.MarginDefaultTopCentimeters);

                document.AddParagraph();

                // Example 2: Using Centimeters (New Method)
                var heading2 = document.AddParagraph("Example 2: Setting Margins Using Centimeters");
                heading2.Bold = true;
                document.AddParagraph("This table uses the new centimeters-based method for setting margins.");

                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Centimeters Example";
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Using New Method";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "All sides 0.2 cm";

                // Set margins using centimeters
                wordTable2.StyleDetails.MarginDefaultTopCentimeters = 0.2;
                wordTable2.StyleDetails.MarginDefaultBottomCentimeters = 0.2;
                wordTable2.StyleDetails.MarginDefaultLeftCentimeters = 0.2;
                wordTable2.StyleDetails.MarginDefaultRightCentimeters = 0.2;
                wordTable2.StyleDetails.CellSpacingCentimeters = 0.1;

                Console.WriteLine("\nSecond table (Centimeters):");
                Console.WriteLine("Table MarginDefaultTopCentimeters: " + wordTable2.StyleDetails.MarginDefaultTopCentimeters);
                Console.WriteLine("Table MarginDefaultTopWidth: " + wordTable2.StyleDetails.MarginDefaultTopWidth);

                document.AddParagraph();

                // Example 3: Mixed Approach with Different Sides
                var heading3 = document.AddParagraph("Example 3: Mixed Approach - Different Units for Different Sides");
                heading3.Bold = true;
                document.AddParagraph("This table demonstrates using both twips and centimeters for different sides.");

                WordTable wordTable3 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                wordTable3.Rows[0].Cells[0].Paragraphs[0].Text = "Mixed Units Example";
                wordTable3.Rows[1].Cells[0].Paragraphs[0].Text = "Using Both Methods";
                wordTable3.Rows[2].Cells[0].Paragraphs[0].Text = "Different sides";

                // Mix of twips and centimeters
                wordTable3.StyleDetails.MarginDefaultTopWidth = 120;  // Using twips
                wordTable3.StyleDetails.MarginDefaultBottomCentimeters = 0.3;  // Using centimeters
                wordTable3.StyleDetails.MarginDefaultLeftWidth = 150;  // Using twips
                wordTable3.StyleDetails.MarginDefaultRightCentimeters = 0.25;  // Using centimeters

                Console.WriteLine("\nThird table (Mixed):");
                Console.WriteLine("Table MarginDefaultTopWidth: " + wordTable3.StyleDetails.MarginDefaultTopWidth);
                Console.WriteLine("Table MarginDefaultBottomCentimeters: " + wordTable3.StyleDetails.MarginDefaultBottomCentimeters);
                Console.WriteLine("Table MarginDefaultLeftWidth: " + wordTable3.StyleDetails.MarginDefaultLeftWidth);
                Console.WriteLine("Table MarginDefaultRightCentimeters: " + wordTable3.StyleDetails.MarginDefaultRightCentimeters);

                document.AddParagraph();

                // Example 4: Border Styles with Both Units
                var heading4 = document.AddParagraph("Example 4: Border Styles with Different Units");
                heading4.Bold = true;
                document.AddParagraph("This table demonstrates border styles with different units for spacing.");

                WordTable wordTable4 = document.AddTable(3, 4, WordTableStyle.GridTable1Light);
                wordTable4.Rows[0].Cells[0].Paragraphs[0].Text = "Border Styles";
                wordTable4.Rows[1].Cells[0].Paragraphs[0].Text = "With Different Units";
                wordTable4.Rows[2].Cells[0].Paragraphs[0].Text = "For Spacing";

                // Set borders with different units
                wordTable4.StyleDetails.SetBordersOutsideInside(
                    BorderValues.Double, 24U, Color.Red,  // Outside borders
                    BorderValues.Single, 12U, Color.Blue  // Inside borders
                );

                // Set cell spacing using centimeters
                wordTable4.StyleDetails.CellSpacingCentimeters = 0.15;

                Console.WriteLine("\nFourth table (Borders):");
                Console.WriteLine("Table CellSpacingCentimeters: " + wordTable4.StyleDetails.CellSpacingCentimeters);
                Console.WriteLine("Table CellSpacing: " + wordTable4.StyleDetails.CellSpacing);

                document.Save(openWord);
            }
        }
    }
}