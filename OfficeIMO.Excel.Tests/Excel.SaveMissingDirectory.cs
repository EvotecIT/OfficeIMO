using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_Save_CreatesMissingDirectory() {
            var sourcePath = Path.Combine(_directoryWithFiles, $"Source_{Guid.NewGuid():N}.xlsx");
            var destinationDirectory = Path.Combine(_directoryWithFiles, "Missing", Guid.NewGuid().ToString("N"));
            var destinationPath = Path.Combine(destinationDirectory, "Created.xlsx");

            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (Directory.Exists(destinationDirectory)) Directory.Delete(destinationDirectory, recursive: true);

            using (var document = ExcelDocument.Create(sourcePath)) {
                const string expectedSheetName = "Sheet1";
                const string expectedCellValue = "Directory save";
                var sheet = document.AddWorkSheet(expectedSheetName);
                sheet.CellValue(1, 1, expectedCellValue);

                document.Save(destinationPath, openExcel: false);

                Assert.True(Directory.Exists(destinationDirectory));
                Assert.True(File.Exists(destinationPath));

                using (var reloaded = ExcelDocument.Load(destinationPath, readOnly: true)) {
                    Assert.Equal(expectedSheetName, reloaded.Sheets[0].Name);
                    Assert.True(reloaded.Sheets[0].TryGetCellText(1, 1, out var actualValue));
                    Assert.Equal(expectedCellValue, actualValue);
                }
            }

            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (File.Exists(destinationPath)) File.Delete(destinationPath);
            if (Directory.Exists(destinationDirectory)) Directory.Delete(destinationDirectory, recursive: true);
        }

        [Fact]
        public async Task Test_SaveAsync_CreatesMissingDirectory() {
            var sourcePath = Path.Combine(_directoryWithFiles, $"AsyncSource_{Guid.NewGuid():N}.xlsx");
            var destinationDirectory = Path.Combine(_directoryWithFiles, "MissingAsync", Guid.NewGuid().ToString("N"));
            var destinationPath = Path.Combine(destinationDirectory, "Created.xlsx");

            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (Directory.Exists(destinationDirectory)) Directory.Delete(destinationDirectory, recursive: true);

            await using (var document = ExcelDocument.Create(sourcePath)) {
                const string expectedSheetName = "AsyncSheet";
                const string expectedCellValue = "Async directory save";
                var sheet = document.AddWorkSheet(expectedSheetName);
                sheet.CellValue(1, 1, expectedCellValue);

                await document.SaveAsync(destinationPath, openExcel: false);

                Assert.True(Directory.Exists(destinationDirectory));
                Assert.True(File.Exists(destinationPath));

                using (var reloaded = ExcelDocument.Load(destinationPath, readOnly: true)) {
                    Assert.Equal(expectedSheetName, reloaded.Sheets[0].Name);
                    Assert.True(reloaded.Sheets[0].TryGetCellText(1, 1, out var actualValue));
                    Assert.Equal(expectedCellValue, actualValue);
                }
            }

            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (File.Exists(destinationPath)) File.Delete(destinationPath);
            if (Directory.Exists(destinationDirectory)) Directory.Delete(destinationDirectory, recursive: true);
        }
    }
}
