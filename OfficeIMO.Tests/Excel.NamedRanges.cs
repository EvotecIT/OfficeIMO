using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_NamedRanges_CreateReadDelete() {
            var filePath = Path.Combine(_directoryWithFiles, "NamedRanges.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet1 = document.AddWorkSheet("Sheet1");
                var sheet2 = document.AddWorkSheet("Sheet2");

                document.CreateNamedRange("GlobalRange", sheet1, "A1:B2", workbookScope: true);
                sheet1.CreateNamedRange("LocalRange", "C1:C3");

                Assert.Equal("Sheet1!A1:B2", document.GetNamedRange("GlobalRange"));
                Assert.Equal("Sheet1!C1:C3", document.GetNamedRange("LocalRange", sheet1));
                Assert.Equal("C1:C3", sheet1.GetNamedRange("LocalRange"));
                Assert.Equal("Sheet1!A1:B2", sheet1.GetNamedRange("GlobalRange", workbookScope: true));

                Assert.True(document.DeleteNamedRange("GlobalRange"));
                Assert.Null(document.GetNamedRange("GlobalRange"));

                Assert.True(sheet1.DeleteNamedRange("LocalRange"));
                Assert.Null(document.GetNamedRange("LocalRange", sheet1));
            }
            File.Delete(filePath);
        }
    }
}
