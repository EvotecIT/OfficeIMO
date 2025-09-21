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
                document.AddWorkSheet("Sheet1");

                document.Save(destinationPath, openExcel: false);

                Assert.True(Directory.Exists(destinationDirectory));
                Assert.True(File.Exists(destinationPath));
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
                document.AddWorkSheet("Sheet1");

                await document.SaveAsync(destinationPath, openExcel: false);

                Assert.True(Directory.Exists(destinationDirectory));
                Assert.True(File.Exists(destinationPath));
            }

            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (File.Exists(destinationPath)) File.Delete(destinationPath);
            if (Directory.Exists(destinationDirectory)) Directory.Delete(destinationDirectory, recursive: true);
        }
    }
}
