using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_BasicExcelExample() {
            var filePath = Path.Combine(_directoryWithFiles, "BasicExcelExample.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                Assert.True(document.Sheets.Count > 0);
            }

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_BasicExcelExampleAsync() {
            var filePath = Path.Combine(_directoryWithFiles, "BasicExcelExampleAsync.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            await using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                await document.SaveAsync();
            }

            await using (var document = await ExcelDocument.LoadAsync(filePath)) {
                Assert.True(document.Sheets.Count > 0);
            }

            File.Delete(filePath);
        }
    }
}