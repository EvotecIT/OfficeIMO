using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private string CreateTestFile(string name) {
            string filePath = Path.Combine(_directoryWithFiles, name);
            if (File.Exists(filePath)) File.Delete(filePath);
            return filePath;
        }

        [Fact]
        public void Test_TableLayoutScenarios() {
            string filePath = CreateTestFile("TestTableLayoutScenarios.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                // Scenario 1: Default Table
                WordTable table1 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table1.Rows[0].Cells[0].Paragraphs[0].Text = "Default";
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, table1.LayoutMode); // Default visually fits window

                // Scenario 2: Full Width (100% Pct) - equivalent to AutoFitToWindow
                WordTable table2 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table2.Rows[0].Cells[0].Paragraphs[0].Text = "100% Pct";
                table2.WidthType = TableWidthUnitValues.Pct;
                table2.Width = 5000;
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, table2.LayoutMode);
                Assert.Equal(TableWidthUnitValues.Pct, table2.WidthType);
                Assert.Equal(5000, table2.Width);

                // Scenario 3: Specific Percentage Width (50%)
                WordTable table3 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table3.Rows[0].Cells[0].Paragraphs[0].Text = "50% Pct";
                table3.WidthType = TableWidthUnitValues.Pct;
                table3.Width = 2500;
                // Setting only width often requires setting LayoutType=Fixed for FixedWidth mode
                table3.LayoutType = TableLayoutValues.Fixed;
                Assert.Equal(WordTableLayoutType.FixedWidth, table3.LayoutMode);
                Assert.Equal(TableLayoutValues.Fixed, table3.LayoutType);
                Assert.Equal(TableWidthUnitValues.Pct, table3.WidthType);
                Assert.Equal(2500, table3.Width);

                // Scenario 4: Using AutoFitToWindow() method
                WordTable table4 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table4.Rows[0].Cells[0].Paragraphs[0].Text = "AutoFit Window Method";
                table4.AutoFitToWindow();
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, table4.LayoutMode);
                Assert.True(table4.LayoutType == TableLayoutValues.Fixed, "Underlying LayoutType should be Fixed after AutoFitToWindow()"); // Explicitly set to Fixed
                Assert.Equal(TableWidthUnitValues.Pct, table4.WidthType);
                Assert.Equal(5000, table4.Width);

                // Scenario 5: Using SetFixedWidth() method (75%)
                WordTable table5 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table5.Rows[0].Cells[0].Paragraphs[0].Text = "Fixed 75% Method";
                table5.SetFixedWidth(75);
                Assert.Equal(WordTableLayoutType.FixedWidth, table5.LayoutMode);
                Assert.Equal(TableLayoutValues.Fixed, table5.LayoutType);
                Assert.Equal(TableWidthUnitValues.Pct, table5.WidthType);
                Assert.Equal(3750, table5.Width);

                // Scenario 6: Using AutoFitToContents() method
                WordTable table6 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table6.Rows[0].Cells[0].Paragraphs[0].Text = "AutoFit Contents Method";
                table6.AutoFitToContents();
                Assert.Equal(WordTableLayoutType.AutoFitToContents, table6.LayoutMode);
                Assert.Equal(TableLayoutValues.Autofit, table6.LayoutType);
                Assert.Equal(TableWidthUnitValues.Auto, table6.WidthType);
                Assert.Equal(0, table6.Width);

                // Scenario 7: Using LayoutMode property setter
                WordTable table7 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table7.Rows[0].Cells[0].Paragraphs[0].Text = "LayoutMode Property";
                table7.LayoutMode = WordTableLayoutType.AutoFitToContents;
                Assert.Equal(WordTableLayoutType.AutoFitToContents, table7.LayoutMode);
                table7.LayoutMode = WordTableLayoutType.AutoFitToWindow;
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, table7.LayoutMode);
                table7.LayoutMode = WordTableLayoutType.FixedWidth; // Sets to 100% Fixed == AutoFitToWindow
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, table7.LayoutMode);
                Assert.Equal(5000, table7.Width);

                document.Save(false);
            }

            // Re-load and verify persistence
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(7, document.Tables.Count);

                // Verify Table 1 (Default)
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, document.Tables[0].LayoutMode);

                // Verify Table 2 (100% Pct)
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, document.Tables[1].LayoutMode);
                Assert.Equal(5000, document.Tables[1].Width);

                // Verify Table 3 (50% Pct Fixed)
                Assert.Equal(WordTableLayoutType.FixedWidth, document.Tables[2].LayoutMode);
                Assert.Equal(TableLayoutValues.Fixed, document.Tables[2].LayoutType);
                Assert.Equal(2500, document.Tables[2].Width);

                // Verify Table 4 (AutoFitToWindow method)
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, document.Tables[3].LayoutMode);
                Assert.Equal(5000, document.Tables[3].Width);

                // Verify Table 5 (SetFixedWidth 75%)
                Assert.Equal(WordTableLayoutType.FixedWidth, document.Tables[4].LayoutMode);
                Assert.Equal(TableLayoutValues.Fixed, document.Tables[4].LayoutType);
                Assert.Equal(3750, document.Tables[4].Width);

                // Verify Table 6 (AutoFitToContents method)
                Assert.Equal(WordTableLayoutType.AutoFitToContents, document.Tables[5].LayoutMode);
                Assert.Equal(TableLayoutValues.Autofit, document.Tables[5].LayoutType);

                // Verify Table 7 (LayoutMode property - ended as Fixed 100% / Window)
                Assert.Equal(WordTableLayoutType.AutoFitToWindow, document.Tables[6].LayoutMode);
                Assert.Equal(5000, document.Tables[6].Width);
            }
        }
    }
}