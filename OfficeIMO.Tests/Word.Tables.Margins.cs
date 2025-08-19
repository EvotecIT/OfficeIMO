using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests table margin properties defined in centimeters.
    /// </summary>
    public partial class Word {
        /// <summary>
        /// Verifies correct handling of table margins specified in centimeters.
        /// </summary>
        [Fact]
        public void Test_TableMarginsWithCentimeters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTableMargins.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                // Test 1: Basic centimeter margins
                WordTable table1 = document.AddTable(3, 3);
                table1.Rows[0].Cells[0].Paragraphs[0].Text = "Centimeter Test";

                var style1 = table1.StyleDetails;
                Assert.NotNull(style1);

                // Set all margins to 0.2 cm
                style1!.MarginDefaultTopCentimeters = 0.2;
                style1.MarginDefaultBottomCentimeters = 0.2;
                style1.MarginDefaultLeftCentimeters = 0.2;
                style1.MarginDefaultRightCentimeters = 0.2;

                // Verify the centimeter values
                Assert.True(Math.Abs(style1.MarginDefaultTopCentimeters.GetValueOrDefault() - 0.2) < 0.01);
                Assert.True(Math.Abs(style1.MarginDefaultBottomCentimeters.GetValueOrDefault() - 0.2) < 0.01);
                Assert.True(Math.Abs(style1.MarginDefaultLeftCentimeters.GetValueOrDefault() - 0.2) < 0.01);
                Assert.True(Math.Abs(style1.MarginDefaultRightCentimeters.GetValueOrDefault() - 0.2) < 0.01);

                // Verify the twips values (0.2 cm should be approximately 113.4 twips)
                Assert.True(Math.Abs(style1.MarginDefaultTopWidth.GetValueOrDefault() - 113) <= 1);
                Assert.True(Math.Abs(style1.MarginDefaultBottomWidth.GetValueOrDefault() - 113) <= 1);
                Assert.True(Math.Abs(style1.MarginDefaultLeftWidth.GetValueOrDefault() - 113) <= 1);
                Assert.True(Math.Abs(style1.MarginDefaultRightWidth.GetValueOrDefault() - 113) <= 1);

                document.AddParagraph();

                // Test 2: Mixed approach (some sides in cm, some in twips)
                WordTable table2 = document.AddTable(3, 3);
                table2.Rows[0].Cells[0].Paragraphs[0].Text = "Mixed Units Test";

                var style2 = table2.StyleDetails;
                Assert.NotNull(style2);

                // Set top and bottom in centimeters, left and right in twips
                style2!.MarginDefaultTopCentimeters = 0.3;
                style2.MarginDefaultBottomCentimeters = 0.3;
                style2.MarginDefaultLeftWidth = 170;
                style2.MarginDefaultRightWidth = 170;

                // Verify centimeter values
                Assert.True(Math.Abs(style2.MarginDefaultTopCentimeters.GetValueOrDefault() - 0.3) < 0.01);
                Assert.True(Math.Abs(style2.MarginDefaultBottomCentimeters.GetValueOrDefault() - 0.3) < 0.01);

                // Verify twips values
                Assert.True(style2.MarginDefaultLeftWidth == 170);
                Assert.True(style2.MarginDefaultRightWidth == 170);

                document.AddParagraph();

                // Test 3: Cell spacing with centimeters
                WordTable table3 = document.AddTable(3, 3);
                table3.Rows[0].Cells[0].Paragraphs[0].Text = "Cell Spacing Test";

                var style3 = table3.StyleDetails;
                Assert.NotNull(style3);

                // Set cell spacing in centimeters
                style3!.CellSpacingCentimeters = 0.15;

                // Verify centimeter value
                Assert.True(Math.Abs(style3.CellSpacingCentimeters.GetValueOrDefault() - 0.15) < 0.01);

                // Verify twips value (0.15 cm should be approximately 85 twips)
                Assert.True(Math.Abs(style3.CellSpacing.GetValueOrDefault() - 85) <= 1);

                document.AddParagraph();

                // Test 4: Null values and clearing
                WordTable table4 = document.AddTable(3, 3);
                table4.Rows[0].Cells[0].Paragraphs[0].Text = "Null Values Test";

                var style4 = table4.StyleDetails;
                Assert.NotNull(style4);

                // Set values and then clear them
                style4!.MarginDefaultTopCentimeters = 0.2;
                style4.MarginDefaultBottomCentimeters = 0.2;
                style4.CellSpacingCentimeters = 0.1;

                // Verify values are set
                Assert.True(Math.Abs(style4.MarginDefaultTopCentimeters.GetValueOrDefault() - 0.2) < 0.01);
                Assert.True(Math.Abs(style4.MarginDefaultBottomCentimeters.GetValueOrDefault() - 0.2) < 0.01);
                Assert.True(Math.Abs(style4.CellSpacingCentimeters.GetValueOrDefault() - 0.1) < 0.01);

                // Clear values
                style4.MarginDefaultTopCentimeters = null;
                style4.MarginDefaultBottomCentimeters = null;
                style4.CellSpacingCentimeters = null;

                // Verify values are cleared
                Assert.True(style4.MarginDefaultTopCentimeters == null);
                Assert.True(style4.MarginDefaultBottomCentimeters == null);
                Assert.True(style4.CellSpacingCentimeters == null);

                document.Save(false);
            }

            // Test 5: Load and verify saved values
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTableMargins.docx"))) {
                // Verify table 1 values
                var table1 = document.Tables[0];
                var styleLoad1 = table1.StyleDetails;
                Assert.NotNull(styleLoad1);
                Assert.True(Math.Abs(styleLoad1!.MarginDefaultTopCentimeters.GetValueOrDefault() - 0.2) < 0.01);
                Assert.True(Math.Abs(styleLoad1.MarginDefaultBottomCentimeters.GetValueOrDefault() - 0.2) < 0.01);
                Assert.True(Math.Abs(styleLoad1.MarginDefaultLeftCentimeters.GetValueOrDefault() - 0.2) < 0.01);
                Assert.True(Math.Abs(styleLoad1.MarginDefaultRightCentimeters.GetValueOrDefault() - 0.2) < 0.01);

                // Verify table 2 values
                var table2 = document.Tables[1];
                var styleLoad2 = table2.StyleDetails;
                Assert.NotNull(styleLoad2);
                Assert.True(Math.Abs(styleLoad2!.MarginDefaultTopCentimeters.GetValueOrDefault() - 0.3) < 0.01);
                Assert.True(Math.Abs(styleLoad2.MarginDefaultBottomCentimeters.GetValueOrDefault() - 0.3) < 0.01);
                Assert.True(styleLoad2.MarginDefaultLeftWidth == 170);
                Assert.True(styleLoad2.MarginDefaultRightWidth == 170);

                // Verify table 3 values
                var table3 = document.Tables[2];
                var styleLoad3 = table3.StyleDetails;
                Assert.NotNull(styleLoad3);
                Assert.True(Math.Abs(styleLoad3!.CellSpacingCentimeters.GetValueOrDefault() - 0.15) < 0.01);

                // Verify table 4 values are cleared
                var table4 = document.Tables[3];
                Assert.True(table4.StyleDetails?.MarginDefaultTopCentimeters == null);
                Assert.True(table4.StyleDetails?.MarginDefaultBottomCentimeters == null);
                Assert.True(table4.StyleDetails?.CellSpacingCentimeters == null);

                document.Save();
            }
        }
    }
}