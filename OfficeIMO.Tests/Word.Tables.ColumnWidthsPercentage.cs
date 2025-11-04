using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void TableColumnWidthsCanBeSetWithPercentages() {
            string filePath = Path.Combine(_directoryWithFiles, "TableColumnWidthsPercentage.docx");

            using (var document = WordDocument.Create(filePath)) {
                var table = document.AddTable(1, 3, WordTableStyle.PlainTable1);
                table.SetColumnWidthsPercentage(10, 30, 60);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                Assert.Equal(5000, table.Width);
                Assert.Equal(TableWidthUnitValues.Pct, table.WidthType);

                Assert.Equal(500, table.Rows[0].Cells[0].Width);
                Assert.Equal(1500, table.Rows[0].Cells[1].Width);
                Assert.Equal(3000, table.Rows[0].Cells[2].Width);

                Assert.Equal(TableWidthUnitValues.Pct, table.Rows[0].Cells[0].WidthType);
                Assert.Equal(TableWidthUnitValues.Pct, table.Rows[0].Cells[1].WidthType);
                Assert.Equal(TableWidthUnitValues.Pct, table.Rows[0].Cells[2].WidthType);
            }
        }

        [Fact]
        public void TableColumnWidthsScalePercentagesWhenTheyDoNotSumToHundred() {
            string filePath = Path.Combine(_directoryWithFiles, "TableColumnWidthsScaledPercentage.docx");

            using (var document = WordDocument.Create(filePath)) {
                var table = document.AddTable(1, 3, WordTableStyle.PlainTable1);
                table.SetColumnWidthsPercentage(1, 1, 8);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                Assert.Equal(500, table.Rows[0].Cells[0].Width);
                Assert.Equal(500, table.Rows[0].Cells[1].Width);
                Assert.Equal(4000, table.Rows[0].Cells[2].Width);
            }
        }

        [Fact]
        public void TableColumnWidthsPercentageRequiresCorrectColumnCount() {
            string filePath = Path.Combine(_directoryWithFiles, "TableColumnWidthsPercentageInvalid.docx");

            using (var document = WordDocument.Create(filePath)) {
                var table = document.AddTable(1, 2, WordTableStyle.PlainTable1);
                Assert.Throws<ArgumentException>(() => table.SetColumnWidthsPercentage(10, 20, 70));
            }
        }

        [Fact]
        public void TableColumnWidthsPercentageSkipsRowsWithFewerCells() {
            string filePath = Path.Combine(_directoryWithFiles, "TableColumnWidthsPercentageMissingCells.docx");

            using (var document = WordDocument.Create(filePath)) {
                var table = document.AddTable(2, 3, WordTableStyle.PlainTable1);
                table.Rows[1].Cells[2].Remove();

                table.SetColumnWidthsPercentage(20, 30, 50);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];

                Assert.Equal(1000, table.Rows[0].Cells[0].Width);
                Assert.Equal(1500, table.Rows[0].Cells[1].Width);
                Assert.Equal(2500, table.Rows[0].Cells[2].Width);

                Assert.Equal(2, table.Rows[1].Cells.Count);
            }
        }
    }
}
