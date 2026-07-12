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

                using (var reloaded = ExcelDocument.Load(destinationPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
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

                using (var reloaded = ExcelDocument.Load(destinationPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
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
        public void Test_SaveCopy_CreatesMissingDirectory() {
            string sourcePath = Path.Combine(_directoryWithFiles, $"CopySource_{Guid.NewGuid():N}.xlsx");
            string destinationDirectory = Path.Combine(_directoryWithFiles, "MissingCopy", Guid.NewGuid().ToString("N"));
            string destinationPath = Path.Combine(destinationDirectory, "Copy.xlsx");

            try {
                using ExcelDocument document = ExcelDocument.Create(sourcePath);
                document.AddWorkSheet("CopyData").CellValue(1, 1, "Directory copy");

                using ExcelDocument copy = document.SaveCopy(destinationPath);

                Assert.True(Directory.Exists(destinationDirectory));
                Assert.True(File.Exists(destinationPath));
                Assert.Equal("CopyData", copy.Sheets[0].Name);
                Assert.True(copy.Sheets[0].TryGetCellText(1, 1, out string? value));
                Assert.Equal("Directory copy", value);
            } finally {
                if (File.Exists(sourcePath)) File.Delete(sourcePath);
                if (Directory.Exists(destinationDirectory)) Directory.Delete(destinationDirectory, recursive: true);
            }
        }

        [Fact]
        public void Test_SaveCopy_PreservesReadOnlyDestination() {
            string sourcePath = Path.Combine(_directoryWithFiles, $"ReadOnlyCopySource_{Guid.NewGuid():N}.xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, $"ReadOnlyCopy_{Guid.NewGuid():N}.xlsx");
            byte[] originalBytes = { 1, 2, 3, 4 };
            File.WriteAllBytes(destinationPath, originalBytes);
            var destination = new FileInfo(destinationPath) { IsReadOnly = true };

            try {
                using ExcelDocument document = ExcelDocument.Create(sourcePath);
                document.AddWorkSheet("Data").CellValue(1, 1, "Must not overwrite");

                Assert.Throws<IOException>(() => document.SaveCopy(destinationPath));

                Assert.Equal(originalBytes, File.ReadAllBytes(destinationPath));
            } finally {
                destination.IsReadOnly = false;
                if (File.Exists(sourcePath)) File.Delete(sourcePath);
                File.Delete(destinationPath);
            }
        }
    }
}
