using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_CellFormattingPersistence() {
            var filePath = Path.Combine(_directoryWithFiles, "FormattingPersistence.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sheet1");
                var cell = sheet.GetCell("B2");
                cell.Text = "Test";
                cell.Font = new Font(new Bold());
                cell.Fill = new Fill(new PatternFill(new ForegroundColor { Rgb = "00FF00" }) { PatternType = PatternValues.Solid });
                cell.Border = new Border(new TopBorder { Style = BorderStyleValues.Thin });
                cell.NumberFormat = "@";
                document.Save();
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets[0];
                var cell = sheet.GetCell("B2");
                Assert.NotNull(cell.Font);
                Assert.NotNull(cell.Fill);
                Assert.NotNull(cell.Border);
                Assert.Equal("@", cell.NumberFormat);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();

            File.Delete(filePath);
        }
    }
}
