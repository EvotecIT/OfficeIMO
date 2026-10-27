using System.Data;
using System.IO;
using System.Threading;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelColumnDefinitionsCancellationTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PreCanceledColumnInspectionPreservesDeferredValues(bool autoFit) {
            using var source = new MemoryStream();
            using var document = ExcelDocument.Create(source);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            ExcelSheet sheet = document.AddWorksheet("Data");
            using DataTable table = Dates();
            sheet.InsertDataTableAsTable(table);
            if (autoFit) sheet.AutoFitColumns();
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
                sheet.GetColumnDefinitions(cancellation.Token));

            Assert.Equal(cancellation.Token, error.CancellationToken);
            AssertSavedDates(document, autoFit);
        }

        [Fact]
        public void ColumnInspectionStopsDeferredMaterializationWhenValueConversionCancels() {
            using var source = new MemoryStream();
            using var document = ExcelDocument.Create(source);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            ExcelSheet sheet = document.AddWorksheet("Data");
            using DataTable table = Dates();
            sheet.InsertDataTableAsTable(table);
            using var cancellation = new CancellationTokenSource();
            int convertedDates = 0;
            document.DateTimeOffsetWriteStrategy = value => {
                convertedDates++;
                cancellation.Cancel();
                return value.UtcDateTime;
            };

            OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
                sheet.GetColumnDefinitions(cancellation.Token));

            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.Equal(1, convertedDates);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            AssertSavedDates(document, autoFit: false);
        }

        private static DataTable Dates() {
            var table = new DataTable("Data");
            table.Columns.Add("When", typeof(DateTimeOffset));
            for (int day = 0; day < 3; day++) table.Rows.Add(Date(day));
            return table;
        }

        private static DateTimeOffset Date(int day) =>
            new DateTimeOffset(2026, 10, 11, 8, 30, 0, TimeSpan.Zero).AddDays(day);

        private static void AssertSavedDates(ExcelDocument document, bool autoFit) {
            using var output = new MemoryStream();
            document.Save(output);
            using var reopened = ExcelDocument.Load(output);
            Assert.Single(reopened.GetTables());
            ExcelSheet sheet = reopened.Sheets[0];
            if (autoFit) Assert.True(Assert.Single(sheet.GetColumnDefinitions()).Width?.CharacterCount > 0D);
            else Assert.Empty(sheet.GetColumnDefinitions());
            for (int day = 0; day < 3; day++) {
                Assert.Equal(Date(day).UtcDateTime.ToOADate(), sheet.CellAt(day + 2, 1).GetValue<double>());
            }
        }
    }
}
