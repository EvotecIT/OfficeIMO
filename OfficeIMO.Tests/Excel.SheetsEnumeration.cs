using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task Test_SheetsEnumerationConcurrent() {
            string filePath = Path.Combine(_directoryWithFiles, "SheetsEnum.xlsx");
            using var document = ExcelDocument.Create(filePath);
            document.AddWorkSheet("S1");
            document.AddWorkSheet("S2");
            document.AddWorkSheet("S3");

            var tasks = Enumerable.Range(0, 10)
                .Select(_ => Task.Run(() => {
                    var sheets = document.Sheets;
                    Assert.Equal(3, sheets.Count);
                }))
                .ToArray();
            await Task.WhenAll(tasks);
            document.Save();
        }
    }
}
