using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void PivotDateFilters_RetainPublishedClrSignatures() {
            var singleDate = new[] { typeof(string), typeof(DateTime), typeof(string), typeof(string) };
            foreach (string name in new[] {
                nameof(ExcelPivotFilter.DateEquals),
                nameof(ExcelPivotFilter.DateNotEquals),
                nameof(ExcelPivotFilter.DateNewerThan),
                nameof(ExcelPivotFilter.DateNewerThanOrEqual),
                nameof(ExcelPivotFilter.DateOlderThan),
                nameof(ExcelPivotFilter.DateOlderThanOrEqual)
            }) {
                Assert.NotNull(typeof(ExcelPivotFilter).GetMethod(name, singleDate));
            }

            var between = new[] { typeof(string), typeof(DateTime), typeof(DateTime), typeof(string), typeof(string) };
            Assert.NotNull(typeof(ExcelPivotFilter).GetMethod(nameof(ExcelPivotFilter.DateBetween), between));
            Assert.NotNull(typeof(ExcelPivotFilter).GetMethod(nameof(ExcelPivotFilter.DateNotBetween), between));

            var generic = new[] { typeof(string), typeof(ExcelPivotFilterType), typeof(DateTime), typeof(DateTime?), typeof(string), typeof(string) };
            Assert.NotNull(typeof(ExcelPivotFilter).GetMethod(nameof(ExcelPivotFilter.Date), generic));
        }
    }
}
