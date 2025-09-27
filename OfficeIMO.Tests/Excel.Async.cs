using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains async Excel tests.
    /// </summary>
    public partial class Excel {
        [Fact]
        public async Task Test_ExcelSaveLoadAsync() {
            var filePath = Path.Combine(_directoryWithFiles, "AsyncExcel.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            await using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                await document.SaveAsync();
            }

            Assert.True(File.Exists(filePath));

            await using (var document = await ExcelDocument.LoadAsync(filePath)) {
                Assert.True(document.Sheets.Count > 0);
            }

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_ExcelLoadAsync_ConcurrentReadWrite() {
            var filePath = Path.Combine(_directoryWithFiles, "AsyncExcelConcurrent.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            await using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                await document.SaveAsync();
            }

            var loadTask1 = ExcelDocument.LoadAsync(filePath, false);
            var loadTask2 = ExcelDocument.LoadAsync(filePath, false);

            var documents = await Task.WhenAll(loadTask1, loadTask2);

            await using var document1 = documents[0];
            await using var document2 = documents[1];
            Assert.True(document1.Sheets.Count > 0);
            Assert.True(document2.Sheets.Count > 0);

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_ExcelSaveAsync_CanBeCancelled() {
            var sourcePath = Path.Combine(_directoryWithFiles, "AsyncExcelCancelSource.xlsx");
            if (File.Exists(sourcePath)) File.Delete(sourcePath);

            var destinationPath = Path.Combine(_directoryWithFiles, "AsyncExcelCancelDestination.xlsx");
            if (File.Exists(destinationPath)) File.Delete(destinationPath);

            await using (var document = ExcelDocument.Create(sourcePath)) {
                document.AddWorkSheet("Sheet1");

                using var cts = new CancellationTokenSource();
                cts.Cancel();

                await Assert.ThrowsAsync<OperationCanceledException>(() => document.SaveAsync(destinationPath, openExcel: false, cancellationToken: cts.Token));
            }

            Assert.False(File.Exists(destinationPath));

            if (File.Exists(sourcePath)) {
                File.Delete(sourcePath);
            }
        }

        [Fact]
        public async Task Test_ExcelSaveAsync_CancellationRestoresWorkbookState() {
            var sourcePath = Path.Combine(_directoryWithFiles, "AsyncExcelCancelRestoreSource.xlsx");
            if (File.Exists(sourcePath)) File.Delete(sourcePath);

            var destinationPath = Path.Combine(_directoryWithFiles, "AsyncExcelCancelRestoreDestination.xlsx");
            if (File.Exists(destinationPath)) File.Delete(destinationPath);

            await using (var document = ExcelDocument.Create(sourcePath)) {
                document.AddWorkSheet("Sheet1");
                var sheet = document.Sheets[0];
                var seedData = Enumerable.Range(0, 200)
                    .Select(i => (Row: i + 1, Column: 1, Value: (object)$"Value {i + 1}"));
                sheet.SetCellValues(seedData);

                using var cts = new CancellationTokenSource();
                var saveTask = document.SaveAsync(destinationPath, openExcel: false, cancellationToken: cts.Token);

                for (int i = 0; i < 200 && !File.Exists(destinationPath); i++) {
                    await Task.Delay(10);
                }

                Assert.True(File.Exists(destinationPath));

                cts.Cancel();

                await Assert.ThrowsAsync<OperationCanceledException>(() => saveTask);

                Assert.NotEmpty(document.Sheets);

                var refreshedSheet = document.Sheets[0];
                refreshedSheet.SetCellValues(new[] { (Row: 1, Column: 2, Value: (object)"Still editable") });

                using var verificationStream = new MemoryStream();
                await document.SaveAsync(verificationStream);
                Assert.True(verificationStream.Length > 0);
            }

            if (File.Exists(destinationPath)) {
                File.Delete(destinationPath);
            }

            if (File.Exists(sourcePath)) {
                File.Delete(sourcePath);
            }
        }
    }
}
