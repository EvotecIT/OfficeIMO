using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelDateTimeOffsetStrategyTests {
        [Theory]
        [InlineData("dom")]
        [InlineData("multipleSheets")]
        [InlineData("direct")]
        [InlineData("outsideA1")]
        [InlineData("overwrite")]
        [InlineData("skipRow")]
        [InlineData("beforeEpoch")]
        public void UpfrontDirectSaveRejectionDoesNotRepeatDateConversion(string route) {
            using var source = new MemoryStream();
            using var document = ExcelDocument.Create(source);
            ExcelSheet sheet = document.AddSheet("Data");
            int row = 1;
            if (route == "dom") {
                sheet.Cell(1, 1).SetValue("Existing\nmultiline cell");
                row = 2;
            } else if (route == "multipleSheets") {
                document.AddSheet("Other");
            } else if (route == "outsideA1") {
                row = 2;
            } else if (route == "overwrite" || route == "skipRow") {
                sheet.Cell(1, 1).SetValue("Existing");
                row = route == "skipRow" ? 3 : 1;
            }
            var date = new DateTimeOffset(route == "beforeEpoch" ? 1800 : 2026, 10, 11, 8, 30, 0, TimeSpan.Zero);
            int conversions = 0;
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime.AddDays(conversions++);

            sheet.Cell(row, 1).SetValue(date);

            Assert.Equal(1, conversions);
            using var saved = new MemoryStream();
            document.Save(saved);
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(1, conversions);
            if (route == "beforeEpoch")
                Assert.Equal(date.ToString("O"), reopened.Sheets[0].Cell(row, 1).GetValue<string>());
            else
                Assert.Equal(date.UtcDateTime.ToOADate(), reopened.Sheets[0].Cell(row, 1).GetValue<double>());
        }
    }
}
