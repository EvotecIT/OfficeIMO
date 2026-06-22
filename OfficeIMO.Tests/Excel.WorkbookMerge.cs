using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelWorkbookMerge_ImportsSelectedSheetsWithPrefix() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.Source.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.Target.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddWorkSheet("North").CellValue(1, 1, "North value");
                source.AddWorkSheet("South").CellValue(1, 1, "South value");
                source.Save();
            }

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, readOnly: true)) {
                target.AddWorkSheet("Summary");
                ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    SheetNames = new[] { "South" },
                    SheetNamePrefix = "Imported "
                });

                Assert.Equal(1, result.SheetCount);
                Assert.Equal(new[] { "South" }, result.SourceSheets);
                Assert.Equal(new[] { "Imported South" }, result.TargetSheets);
                Assert.True(target["Imported South"].TryGetCellText(1, 1, out var importedValue));
                Assert.Equal("South value", importedValue);
                target.Save();
            }

            using (var reloaded = ExcelDocument.Load(targetPath, readOnly: true)) {
                Assert.True(reloaded["Imported South"].TryGetCellText(1, 1, out var importedValue));
                Assert.Equal("South value", importedValue);
            }
        }
    }
}
