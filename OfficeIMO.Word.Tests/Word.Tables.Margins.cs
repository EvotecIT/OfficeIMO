using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;
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

                // Set all margins to 0.2 cm
                table1.StyleDetails!.MarginDefaultTop = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                table1.StyleDetails!.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                table1.StyleDetails!.MarginDefaultLeft = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                table1.StyleDetails!.MarginDefaultRight = global::OfficeIMO.OfficeLength.Centimeters(0.2);

                // Verify the centimeter values
                Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultTop.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultBottom.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultLeft.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);
                Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultRight.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);

                // Verify the twips values (0.2 cm should be approximately 113.4 twips)
                Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultTop.GetValueOrDefault().ToTwips() - 113) <= 1);
                Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultBottom.GetValueOrDefault().ToTwips() - 113) <= 1);
                Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultLeft.GetValueOrDefault().ToTwips() - 113) <= 1);
                Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultRight.GetValueOrDefault().ToTwips() - 113) <= 1);

                document.AddParagraph();

                // Test 2: Mixed approach (some sides in cm, some in twips)
                WordTable table2 = document.AddTable(3, 3);
                table2.Rows[0].Cells[0].Paragraphs[0].Text = "Mixed Units Test";

                // Set top and bottom in centimeters, left and right in twips
                table2.StyleDetails!.MarginDefaultTop = global::OfficeIMO.OfficeLength.Centimeters(0.3);
                table2.StyleDetails!.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Centimeters(0.3);
                table2.StyleDetails!.MarginDefaultLeft = global::OfficeIMO.OfficeLength.Points(8.5);
                table2.StyleDetails!.MarginDefaultRight = global::OfficeIMO.OfficeLength.Points(8.5);

                // Verify centimeter values
                Assert.True(Math.Abs(table2.StyleDetails!.MarginDefaultTop.GetValueOrDefault().ToCentimeters() - 0.3) < 0.01);
                Assert.True(Math.Abs(table2.StyleDetails!.MarginDefaultBottom.GetValueOrDefault().ToCentimeters() - 0.3) < 0.01);

                // Verify twips values
                Assert.True(table2.StyleDetails!.MarginDefaultLeft == global::OfficeIMO.OfficeLength.Points(8.5));
                Assert.True(table2.StyleDetails!.MarginDefaultRight == global::OfficeIMO.OfficeLength.Points(8.5));

                document.AddParagraph();

                // Test 3: Cell spacing with centimeters
                WordTable table3 = document.AddTable(3, 3);
                table3.Rows[0].Cells[0].Paragraphs[0].Text = "Cell Spacing Test";

                // Set cell spacing in centimeters
                table3.StyleDetails!.CellSpacing = global::OfficeIMO.OfficeLength.Centimeters(0.15);

                // Verify centimeter value
                Assert.True(Math.Abs(table3.StyleDetails!.CellSpacing.GetValueOrDefault().ToCentimeters() - 0.15) < 0.01);

                // Verify twips value (0.15 cm should be approximately 85 twips)
                Assert.True(Math.Abs(table3.StyleDetails!.CellSpacing.GetValueOrDefault().ToTwips() - 85) <= 1);

                document.AddParagraph();

                // Test 4: Null values and clearing
                WordTable table4 = document.AddTable(3, 3);
                table4.Rows[0].Cells[0].Paragraphs[0].Text = "Null Values Test";

                // Set values and then clear them
                table4.StyleDetails!.MarginDefaultTop = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                table4.StyleDetails!.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Centimeters(0.2);
                table4.StyleDetails!.CellSpacing = global::OfficeIMO.OfficeLength.Centimeters(0.1);

                // Verify values are set
                  Assert.True(Math.Abs(table4.StyleDetails!.MarginDefaultTop.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);
                  Assert.True(Math.Abs(table4.StyleDetails!.MarginDefaultBottom.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);
                  Assert.True(Math.Abs(table4.StyleDetails!.CellSpacing.GetValueOrDefault().ToCentimeters() - 0.1) < 0.01);

                // Clear values
                table4.StyleDetails!.MarginDefaultTop = null;
                table4.StyleDetails!.MarginDefaultBottom = null;
                table4.StyleDetails!.CellSpacing = null;

                // Verify values are cleared
                Assert.True(table4.StyleDetails!.MarginDefaultTop == null);
                Assert.True(table4.StyleDetails!.MarginDefaultBottom == null);
                Assert.True(table4.StyleDetails!.CellSpacing == null);

                document.Save();
            }

            // Test 5: Load and verify saved values
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTableMargins.docx"))) {
                // Verify table 1 values
                var table1 = document.Tables[0];
                  Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultTop.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);
                  Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultBottom.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);
                  Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultLeft.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);
                  Assert.True(Math.Abs(table1.StyleDetails!.MarginDefaultRight.GetValueOrDefault().ToCentimeters() - 0.2) < 0.01);

                // Verify table 2 values
                var table2 = document.Tables[1];
                  Assert.True(Math.Abs(table2.StyleDetails!.MarginDefaultTop.GetValueOrDefault().ToCentimeters() - 0.3) < 0.01);
                  Assert.True(Math.Abs(table2.StyleDetails!.MarginDefaultBottom.GetValueOrDefault().ToCentimeters() - 0.3) < 0.01);
                Assert.True(table2.StyleDetails!.MarginDefaultLeft == global::OfficeIMO.OfficeLength.Points(8.5));
                Assert.True(table2.StyleDetails!.MarginDefaultRight == global::OfficeIMO.OfficeLength.Points(8.5));

                // Verify table 3 values
                var table3 = document.Tables[2];
                  Assert.True(Math.Abs(table3.StyleDetails!.CellSpacing.GetValueOrDefault().ToCentimeters() - 0.15) < 0.01);

                // Verify table 4 values are cleared
                var table4 = document.Tables[3];
                Assert.True(table4.StyleDetails!.MarginDefaultTop == null);
                Assert.True(table4.StyleDetails!.MarginDefaultBottom == null);
                Assert.True(table4.StyleDetails!.CellSpacing == null);

                document.Save();
            }
        }
    }
}