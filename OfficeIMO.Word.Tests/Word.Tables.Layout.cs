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
                Assert.Equal(WordTableLayoutMode.AutoFit, table1.LayoutMode);

                // Scenario 2: Full Width (100% Pct) - equivalent to AutoFitToWindow
                WordTable table2 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table2.Rows[0].Cells[0].Paragraphs[0].Text = "100% Pct";
                table2.Width = WordTableWidth.Percent((5000) / 50d);
                Assert.Equal(WordTableLayoutMode.AutoFit, table2.LayoutMode);
                Assert.Equal(WordTableWidthKind.Percent, table2.Width?.Kind);
                Assert.Equal(5000, table2.Width?.ToNativeValue());

                // Scenario 3: Specific Percentage Width (50%)
                WordTable table3 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table3.Rows[0].Cells[0].Paragraphs[0].Text = "50% Pct";
                table3.Width = WordTableWidth.Percent((2500) / 50d);
                // A percentage below 100% is a fixed-width table.
                table3.LayoutMode = WordTableLayoutMode.Fixed;
                Assert.Equal(WordTableLayoutMode.Fixed, table3.LayoutMode);
                Assert.Equal(WordTableWidthKind.Percent, table3.Width?.Kind);
                Assert.Equal(2500, table3.Width?.ToNativeValue());

                // Scenario 4: Using AutoFitToWindow() method
                WordTable table4 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table4.Rows[0].Cells[0].Paragraphs[0].Text = "AutoFit Window Method";
                table4.AutoFitToWindow();
                Assert.Equal(WordTableLayoutMode.Fixed, table4.LayoutMode);
                Assert.Equal(WordTableWidthKind.Percent, table4.Width?.Kind);
                Assert.Equal(5000, table4.Width?.ToNativeValue());

                // Scenario 5: Using SetFixedWidth() method (75%)
                WordTable table5 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table5.Rows[0].Cells[0].Paragraphs[0].Text = "Fixed 75% Method";
                table5.SetFixedWidth(75);
                Assert.Equal(WordTableLayoutMode.Fixed, table5.LayoutMode);
                Assert.Equal(WordTableWidthKind.Percent, table5.Width?.Kind);
                Assert.Equal(3750, table5.Width?.ToNativeValue());

                // Scenario 6: Using AutoFitToContents() method
                WordTable table6 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table6.Rows[0].Cells[0].Paragraphs[0].Text = "AutoFit Contents Method";
                table6.AutoFitToContents();
                Assert.Equal(WordTableLayoutMode.AutoFit, table6.LayoutMode);
                Assert.Equal(WordTableWidthKind.Auto, table6.Width?.Kind);
                Assert.Equal(0, table6.Width?.ToNativeValue());

                // Scenario 7: Using LayoutMode property setter
                WordTable table7 = document.AddTable(3, 3, WordTableStyle.PlainTable1);
                table7.Rows[0].Cells[0].Paragraphs[0].Text = "LayoutMode Property";
                table7.LayoutMode = WordTableLayoutMode.AutoFit;
                Assert.Equal(WordTableLayoutMode.AutoFit, table7.LayoutMode);
                table7.AutoFitToWindow();
                Assert.Equal(WordTableLayoutMode.Fixed, table7.LayoutMode);
                table7.LayoutMode = WordTableLayoutMode.Fixed;
                Assert.Equal(WordTableLayoutMode.Fixed, table7.LayoutMode);
                Assert.Equal(5000, table7.Width?.ToNativeValue());

                document.Save();
            }

            // Re-load and verify persistence
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(7, document.Tables.Count);

                // Verify Table 1 (Default)
                Assert.Equal(WordTableLayoutMode.AutoFit, document.Tables[0].LayoutMode);

                // Verify Table 2 (100% Pct)
                Assert.Equal(WordTableLayoutMode.AutoFit, document.Tables[1].LayoutMode);
                Assert.Equal(5000, document.Tables[1].Width?.ToNativeValue());

                // Verify Table 3 (50% Pct Fixed)
                Assert.Equal(WordTableLayoutMode.Fixed, document.Tables[2].LayoutMode);
                Assert.Equal(2500, document.Tables[2].Width?.ToNativeValue());

                // Verify Table 4 (AutoFitToWindow method)
                Assert.Equal(WordTableLayoutMode.Fixed, document.Tables[3].LayoutMode);
                Assert.Equal(5000, document.Tables[3].Width?.ToNativeValue());

                // Verify Table 5 (SetFixedWidth 75%)
                Assert.Equal(WordTableLayoutMode.Fixed, document.Tables[4].LayoutMode);
                Assert.Equal(3750, document.Tables[4].Width?.ToNativeValue());

                // Verify Table 6 (AutoFitToContents method)
                Assert.Equal(WordTableLayoutMode.AutoFit, document.Tables[5].LayoutMode);

                // Verify Table 7 (AutoFitToWindow helper followed by the fixed algorithm)
                Assert.Equal(WordTableLayoutMode.Fixed, document.Tables[6].LayoutMode);
                Assert.Equal(5000, document.Tables[6].Width?.ToNativeValue());
            }
        }

        [Fact]
        public void LayoutMode_ChangesAlgorithmWithoutRewritingPreferredWidths() {
            string filePath = CreateTestFile("TableLayoutPreservesWidths.docx");

            using WordDocument document = WordDocument.Create(filePath);
            WordTable table = document.AddTable(1, 2);
            table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(4320) * 635)));
            table.Rows[0].Cells[0].Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(1800) * 635)));

            table.LayoutMode = WordTableLayoutMode.AutoFit;
            Assert.Equal(WordTableWidthKind.Length, table.Width?.Kind);
            Assert.Equal(4320, table.Width?.ToNativeValue());
            Assert.Equal(1800, table.Rows[0].Cells[0].Width?.ToNativeValue());

            table.LayoutMode = WordTableLayoutMode.Fixed;
            Assert.Equal(WordTableWidthKind.Length, table.Width?.Kind);
            Assert.Equal(4320, table.Width?.ToNativeValue());
            Assert.Equal(1800, table.Rows[0].Cells[0].Width?.ToNativeValue());
        }
    }
}
