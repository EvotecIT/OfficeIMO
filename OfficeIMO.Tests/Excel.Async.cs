using System.IO;
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
    }
}
