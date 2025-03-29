using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_TableMarginsWithCentimeters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTableMargins.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                // Test 1: Basic centimeter margins
                WordTable table1 = document.AddTable(3, 3);
                table1.Rows[0].Cells[0].Paragraphs[0].Text = "Centimeter Test";

                // Set all margins to 0.2 cm
                table1.StyleDetails.MarginDefaultTopCentimeters = 0.2;
                table1.StyleDetails.MarginDefaultBottomCentimeters = 0.2;
                table1.StyleDetails.MarginDefaultLeftCentimeters = 0.2;
                table1.StyleDetails.MarginDefaultRightCentimeters = 0.2;

                // Verify the centimeter values
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultTopCentimeters.Value - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultBottomCentimeters.Value - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultLeftCentimeters.Value - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultRightCentimeters.Value - 0.2) < 0.01);

                // Verify the twips values (0.2 cm should be approximately 113.4 twips)
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultTopWidth.Value - 113) <= 1);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultBottomWidth.Value - 113) <= 1);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultLeftWidth.Value - 113) <= 1);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultRightWidth.Value - 113) <= 1);

                document.AddParagraph();

                // Test 2: Mixed approach (some sides in cm, some in twips)
                WordTable table2 = document.AddTable(3, 3);
                table2.Rows[0].Cells[0].Paragraphs[0].Text = "Mixed Units Test";

                // Set top and bottom in centimeters, left and right in twips
                table2.StyleDetails.MarginDefaultTopCentimeters = 0.3;
                table2.StyleDetails.MarginDefaultBottomCentimeters = 0.3;
                table2.StyleDetails.MarginDefaultLeftWidth = 170;
                table2.StyleDetails.MarginDefaultRightWidth = 170;

                // Verify centimeter values
                Assert.True(Math.Abs(table2.StyleDetails.MarginDefaultTopCentimeters.Value - 0.3) < 0.01);
                Assert.True(Math.Abs(table2.StyleDetails.MarginDefaultBottomCentimeters.Value - 0.3) < 0.01);

                // Verify twips values
                Assert.True(table2.StyleDetails.MarginDefaultLeftWidth == 170);
                Assert.True(table2.StyleDetails.MarginDefaultRightWidth == 170);

                document.AddParagraph();

                // Test 3: Cell spacing with centimeters
                WordTable table3 = document.AddTable(3, 3);
                table3.Rows[0].Cells[0].Paragraphs[0].Text = "Cell Spacing Test";

                // Set cell spacing in centimeters
                table3.StyleDetails.CellSpacingCentimeters = 0.15;

                // Verify centimeter value
                Assert.True(Math.Abs(table3.StyleDetails.CellSpacingCentimeters.Value - 0.15) < 0.01);

                // Verify twips value (0.15 cm should be approximately 85 twips)
                Assert.True(Math.Abs(table3.StyleDetails.CellSpacing.Value - 85) <= 1);

                document.AddParagraph();

                // Test 4: Null values and clearing
                WordTable table4 = document.AddTable(3, 3);
                table4.Rows[0].Cells[0].Paragraphs[0].Text = "Null Values Test";

                // Set values and then clear them
                table4.StyleDetails.MarginDefaultTopCentimeters = 0.2;
                table4.StyleDetails.MarginDefaultBottomCentimeters = 0.2;
                table4.StyleDetails.CellSpacingCentimeters = 0.1;

                // Verify values are set
                Assert.True(Math.Abs(table4.StyleDetails.MarginDefaultTopCentimeters.Value - 0.2) < 0.01);
                Assert.True(Math.Abs(table4.StyleDetails.MarginDefaultBottomCentimeters.Value - 0.2) < 0.01);
                Assert.True(Math.Abs(table4.StyleDetails.CellSpacingCentimeters.Value - 0.1) < 0.01);

                // Clear values
                table4.StyleDetails.MarginDefaultTopCentimeters = null;
                table4.StyleDetails.MarginDefaultBottomCentimeters = null;
                table4.StyleDetails.CellSpacingCentimeters = null;

                // Verify values are cleared
                Assert.True(table4.StyleDetails.MarginDefaultTopCentimeters == null);
                Assert.True(table4.StyleDetails.MarginDefaultBottomCentimeters == null);
                Assert.True(table4.StyleDetails.CellSpacingCentimeters == null);

                document.Save(false);
            }

            // Test 5: Load and verify saved values
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTableMargins.docx"))) {
                // Verify table 1 values
                var table1 = document.Tables[0];
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultTopCentimeters.Value - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultBottomCentimeters.Value - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultLeftCentimeters.Value - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails.MarginDefaultRightCentimeters.Value - 0.2) < 0.01);

                // Verify table 2 values
                var table2 = document.Tables[1];
                Assert.True(Math.Abs(table2.StyleDetails.MarginDefaultTopCentimeters.Value - 0.3) < 0.01);
                Assert.True(Math.Abs(table2.StyleDetails.MarginDefaultBottomCentimeters.Value - 0.3) < 0.01);
                Assert.True(table2.StyleDetails.MarginDefaultLeftWidth == 170);
                Assert.True(table2.StyleDetails.MarginDefaultRightWidth == 170);

                // Verify table 3 values
                var table3 = document.Tables[2];
                Assert.True(Math.Abs(table3.StyleDetails.CellSpacingCentimeters.Value - 0.15) < 0.01);

                // Verify table 4 values are cleared
                var table4 = document.Tables[3];
                Assert.True(table4.StyleDetails.MarginDefaultTopCentimeters == null);
                Assert.True(table4.StyleDetails.MarginDefaultBottomCentimeters == null);
                Assert.True(table4.StyleDetails.CellSpacingCentimeters == null);

                document.Save();
            }
        }
    }
}