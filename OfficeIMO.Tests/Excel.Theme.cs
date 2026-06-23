using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelWorkbookTheme_CanResetRenameAndReadBack() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookTheme.ResetRename.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data");
                document.ResetWorkbookTheme("Contoso Workbook Theme");
                ExcelWorkbookThemeInfo info = document.GetWorkbookTheme(includeXml: true);

                Assert.True(info.HasTheme);
                Assert.Equal("Contoso Workbook Theme", info.Name);
                Assert.Contains("Contoso Workbook Theme", info.Xml);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ThemePart themePart = spreadsheet.WorkbookPart!.GetPartsOfType<ThemePart>().Single();
                using var reader = new StreamReader(themePart.GetStream(FileMode.Open, FileAccess.Read));
                string xml = reader.ReadToEnd();
                Assert.Contains("Contoso Workbook Theme", xml);
            }
        }
    }
}
