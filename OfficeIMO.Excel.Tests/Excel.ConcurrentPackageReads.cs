using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelConcurrentPackageReads {
        // Each read loads pivot, pivot cache and style parts the package hasn't loaded yet; concurrent loads of one
        // package corrupt each other unless reads take turns.
        [Fact]
        public async Task ConcurrentReadsOfAnOpenPackageAllSucceed() {
            byte[] workbook = PivotWorkbook();
            for (int round = 0; round < 20; round++) {
                using var document = ExcelDocument.Load(new MemoryStream(workbook, false));
                _ = document.OpenXmlDocument;
                using var start = new ManualResetEventSlim(false);
                Task<int>[] readers = Enumerable.Range(0, 8).Select(_ => Task.Run(() => {
                    start.Wait();
                    return document.GetPivotTables().Count;
                })).ToArray();
                start.Set();

                Assert.All(await Task.WhenAll(readers), count => Assert.Equal(1, count));
            }
        }

        private static byte[] PivotWorkbook() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet data = document.AddSheet("Data");
                data.Cell(1, 1).SetValue("Region");
                data.Cell(1, 2).SetValue("Sales");
                data.Cell(2, 1).SetValue("East");
                data.Cell(2, 2).SetValue(10d);
                data.AddPivotTable("A1:B2", "D2", options: new ExcelPivotTableOptions {
                    Name = "SalesPivot",
                    RowFields = new[] { "Region" },
                    DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) }
                });
                document.Save();
            }
            return stream.ToArray();
        }
    }
}
