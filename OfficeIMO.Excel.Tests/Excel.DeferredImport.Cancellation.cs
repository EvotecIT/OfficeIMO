using System.Data;
using System.IO;
using System.Threading;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelDeferredImportCancellationTests {
        [Fact]
        public void PreCanceledWorksheetCreationPreservesTheExistingDeferredImport() {
            using var source = new MemoryStream();
            using var document = ExcelDocument.Create(source);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            ExcelSheet sheet = document.AddSheet("First");
            using DataTable table = Dates();
            sheet.InsertDataTableAsTable(table);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
                document.AddSheet("Second", ExcelSheetNameValidationMode.Strict, cancellation.Token));

            Assert.Equal(cancellation.Token, error.CancellationToken);
            AssertSavedDates(document);
        }

        [Fact]
        public void WorksheetCreationStopsPriorDeferredMaterializationWhenValueConversionCancels() {
            using var source = new MemoryStream();
            using var document = ExcelDocument.Create(source);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            ExcelSheet sheet = document.AddSheet("First");
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
                document.AddSheet("Second", ExcelSheetNameValidationMode.Strict, cancellation.Token));

            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.Equal(1, convertedDates);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            AssertSavedDates(document);
        }

        [Theory]
        [InlineData("replace")]
        [InlineData("append")]
        [InlineData("dataset")]
        [InlineData("replace-table")]
        [InlineData("append-table")]
        public void ImportStopsPriorDeferredMaterializationWhenValueConversionCancels(string route) {
            using var source = new MemoryStream();
            using var document = ExcelDocument.Create(source);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            ExcelSheet sheet = document.AddSheet("First");
            using DataTable prior = Dates();
            sheet.InsertDataTableAsTable(prior, tableName: "Existing");
            using var next = new DataTable("Replacement");
            next.Columns.Add("Value", typeof(string));
            next.Rows.Add("New value");
            using var cancellation = new CancellationTokenSource();
            int convertedDates = 0;
            document.DateTimeOffsetWriteStrategy = value => {
                convertedDates++;
                cancellation.Cancel();
                return value.UtcDateTime;
            };
            using var dataSet = new DataSet();
            if (route == "dataset") dataSet.Tables.Add(next);

            OperationCanceledException error = Assert.Throws<OperationCanceledException>(() => {
                if (route == "dataset") document.InsertDataSet(dataSet, ct: cancellation.Token);
                else if (route == "replace-table") sheet.InsertDataTableAsTable(next, ct: cancellation.Token);
                else if (route == "append-table") sheet.AppendDataTableToTable(next, "Existing", ct: cancellation.Token);
                else sheet.WriteData(next, row: route == "append" ? 10 : 1, cancellationToken: cancellation.Token);
            });

            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.Equal(1, convertedDates);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            AssertSavedDates(document);
        }

        [Fact]
        public void ReaderDomFallbackStopsBeforeWritingLaterCellsAfterCancellation() {
            using var source = new MemoryStream();
            using var document = ExcelDocument.Create(source);
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            ExcelSheet sheet = document.AddSheet("First");
            using var prior = new DataTable("Prior");
            prior.Columns.Add("Original", typeof(string));
            prior.Rows.Add("Existing deferred value");
            sheet.WriteData(prior);
            using DataTable table = Dates();
            table.Columns.Add("Description", typeof(string));
            for (int row = 0; row < table.Rows.Count; row++) table.Rows[row][1] = "Later cell " + row;
            using var cancellation = new CancellationTokenSource();
            var convertedDates = new List<DateTimeOffset>();
            document.DateTimeOffsetWriteStrategy = value => {
                convertedDates.Add(value);
                cancellation.Cancel();
                return value.UtcDateTime;
            };
            using var reader = table.CreateDataReader();

            OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
                sheet.InsertDataReaderWithResult(reader, createTable: false, autoFit: false, ct: cancellation.Token));

            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.True(string.IsNullOrEmpty(sheet.CellAt(2, 2).GetValue<string>()));
            Assert.True(string.IsNullOrEmpty(sheet.CellAt(3, 1).GetValue<string>()));
            Assert.NotEmpty(convertedDates);
            Assert.All(convertedDates, value => Assert.Equal(Date(0), value));
            document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
            using var saved = new MemoryStream();
            document.Save(saved);
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(Date(0).UtcDateTime.ToOADate(), reopened.Sheets[0].CellAt(2, 1).GetValue<double>());
            Assert.True(string.IsNullOrEmpty(reopened.Sheets[0].CellAt(2, 2).GetValue<string>()));
            Assert.True(string.IsNullOrEmpty(reopened.Sheets[0].CellAt(3, 1).GetValue<string>()));
        }

        private static DataTable Dates() {
            var table = new DataTable("Data");
            table.Columns.Add("When", typeof(DateTimeOffset));
            for (int day = 0; day < 3; day++) table.Rows.Add(Date(day));
            return table;
        }

        private static DateTimeOffset Date(int day) =>
            new DateTimeOffset(2026, 10, 11, 8, 30, 0, TimeSpan.Zero).AddDays(day);

        private static void AssertSavedDates(ExcelDocument document) {
            using var output = new MemoryStream();
            document.Save(output);
            using var reopened = ExcelDocument.Load(output);
            ExcelSheet sheet = Assert.Single(reopened.Sheets);
            Assert.Equal("First", sheet.Name);
            Assert.Single(reopened.GetTables());
            for (int day = 0; day < 3; day++) {
                Assert.Equal(Date(day).UtcDateTime.ToOADate(), sheet.CellAt(day + 2, 1).GetValue<double>());
            }
        }
    }
}
