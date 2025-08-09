using System.IO;
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

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                await document.SaveAsync();
            }

            Assert.True(File.Exists(filePath));

            using (var document = await ExcelDocument.LoadAsync(filePath)) {
                Assert.True(document.Sheets.Count > 0);
            }

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_ExcelSaveAsyncCancelledDoesNotWriteFile() {
            var sourcePath = Path.Combine(_directoryWithFiles, "AsyncSource.xlsx");
            var targetPath = Path.Combine(_directoryWithFiles, "AsyncCancelled.xlsx");
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (File.Exists(targetPath)) File.Delete(targetPath);

            using (var document = ExcelDocument.Create(sourcePath)) {
                document.AddWorkSheet("Sheet1");
                using var cts = new CancellationTokenSource();
                cts.Cancel();
                await Assert.ThrowsAsync<OperationCanceledException>(() => document.SaveAsync(targetPath, false, cts.Token));
            }

            Assert.False(File.Exists(targetPath));
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
        }
    }
}
