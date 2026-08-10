using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_DirectTabularSave_WritesSupportedScalarValuesThroughPublicApi() {
            var timings = new List<string>();
            using ExcelDocument document = ExcelDocument.Create();
            document.Execution.OnTiming = (operation, _) => timings.Add(operation);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Text");
            sheet.CellValue(1, 2, "Integer");
            sheet.CellValue(1, 3, "Decimal");
            sheet.CellValue(1, 4, "Boolean");
            sheet.CellValue(2, 1, "Zażółć gęślą");
            sheet.CellValue(2, 2, 42);
            sheet.CellValue(2, 3, 12.5m);
            sheet.CellValue(2, 4, true);
            for (int row = 3; row <= 256; row++) {
                sheet.CellValue(row, 1, "Row " + row);
                sheet.CellValue(row, 2, row);
                sheet.CellValue(row, 3, row + 0.5m);
                sheet.CellValue(row, 4, (row & 1) == 0);
            }

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xls);

            Assert.Contains("Save.Xls.Direct.ExtractCells", timings);
            Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.Equal(4, reader.FieldCount);
            Assert.Equal("Text", reader.GetName(0));
            Assert.Equal("Integer", reader.GetName(1));
            Assert.Equal("Decimal", reader.GetName(2));
            Assert.Equal("Boolean", reader.GetName(3));
            Assert.True(reader.Read());
            Assert.Equal("Zażółć gęślą", reader.GetString(0));
            Assert.Equal(42, reader.GetInt32(1));
            Assert.Equal(12.5m, reader.GetDecimal(2));
            Assert.True(reader.GetBoolean(3));
            int rowCount = 1;
            while (reader.Read()) {
                rowCount++;
            }
            Assert.Equal(255, rowCount);
            Assert.False(reader.NextResult());
        }

        [Fact]
        public void LegacyXls_DirectTabularSave_FallsBackForDateValuesWithoutChangingSemantics() {
            DateTime expected = new DateTime(2026, 8, 10, 14, 30, 0, DateTimeKind.Unspecified);
            var messages = new List<string>();
            using ExcelDocument document = ExcelDocument.Create();
            document.Execution.OnInfo = messages.Add;
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "When");
            sheet.CellValue(2, 1, expected);

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xls);

            Assert.Contains(messages, message => message.Contains("requires materialization or preflight", StringComparison.Ordinal));
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.True(reader.Read());
            Assert.Equal(expected, reader.GetDateTime(0));
            Assert.False(reader.Read());
        }

        [Fact]
        public async Task LegacyXls_DirectTabularSave_CoversSyncAndAsyncFilePathsAndRemainsMutable() {
            string syncPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            string asyncPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            try {
                var timings = new List<string>();
                using ExcelDocument document = ExcelDocument.Create();
                document.Execution.OnTiming = (operation, _) => timings.Add(operation);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                for (int row = 2; row <= 256; row++) {
                    sheet.CellValue(row, 1, row - 1);
                }

                document.Save(syncPath);
                Assert.Contains("Save.Xls.Direct.ExtractCells", timings);
                Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
                AssertSingleValue(syncPath, 1);
                AssertWorkbookOpensViaExcelComWhenAvailable(
                    syncPath,
                    "The directly generated XLS workbook failed to open in desktop Excel.");

                timings.Clear();
                sheet.CellValue(2, 1, 2);
                await document.SaveAsync(asyncPath);
                AssertSingleValue(asyncPath, 2);
            } finally {
                TryDelete(syncPath);
                TryDelete(asyncPath);
            }
        }

        [Fact]
        public void LegacyXls_DirectTabularSave_PreservesHeaderlessAllNullShape() {
            var table = new DataTable("Data");
            table.Columns.Add("A", typeof(string));
            table.Columns.Add("B", typeof(string));
            table.Columns.Add("C", typeof(string));
            for (int row = 0; row < 65; row++) {
                table.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);
            }
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            using ExcelDocument document = ExcelDocument.Create();
            document.InsertDataSet(
                dataSet,
                createTables: false,
                includeHeaders: false,
                includeAutoFilter: false);
            Assert.True(document.HasDeferredDirectDataSetImport);

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xls);

            Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            Assert.True(document.HasDeferredDirectDataSetImport);
            AssertBiffIndexMatchesDimensions(
                workbook,
                expectedDataFirstRow: 0,
                expectedDataRowAfterLast: 0,
                expectedDimensionFirstRow: 0,
                expectedDimensionRowAfterLast: 65);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
                workbook,
                new ExcelReadOptions { HasHeaderRow = false });
            Assert.Equal(3, reader.FieldCount);
            int rowCount = 0;
            while (reader.Read()) {
                Assert.True(reader.IsDBNull(0));
                Assert.True(reader.IsDBNull(1));
                Assert.True(reader.IsDBNull(2));
                rowCount++;
            }
            Assert.Equal(65, rowCount);

            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            try {
                File.WriteAllBytes(path, workbook);
                AssertWorkbookOpensViaExcelComWhenAvailable(
                    path,
                    "The direct BIFF8 workbook with empty coordinate bands failed to open in desktop Excel.");
            } finally {
                TryDelete(path);
            }
        }

        [Fact]
        public void LegacyXls_DirectTabularSave_PreservesTrailingBlankRowAndColumn() {
            var table = new DataTable("Data");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Unused", typeof(string));
            table.Rows.Add(1, "Alpha", DBNull.Value);
            table.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            using ExcelDocument document = ExcelDocument.Create();
            document.InsertDataSet(
                dataSet,
                createTables: false,
                includeAutoFilter: false);

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xls);

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.Equal(3, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.Equal("Alpha", reader.GetString(1));
            Assert.True(reader.IsDBNull(2));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.IsDBNull(1));
            Assert.True(reader.IsDBNull(2));
            Assert.False(reader.Read());
        }

        [Fact]
        public async Task LegacyXls_DirectTabularSave_HonorsPreCanceledAsyncStreamWrite() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            sheet.CellValue(2, 1, 1);
            using var destination = new MemoryStream();
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                document.SaveAsync(destination, ExcelFileFormat.Xls, cancellationToken: cancellation.Token));
            Assert.Equal(0, destination.Length);
        }

        [Fact]
        public void LegacyXls_GeneralWriter_OpensInDesktopExcelAndRoundTripsValues() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            try {
                using ExcelDocument document = ExcelDocument.Create();
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                for (int row = 2; row <= 256; row++) {
                    sheet.CellValue(row, 1, row - 1);
                }

                document.Save(path, new ExcelSaveOptions { DisableFastPackageWriter = true });

                Assert.NotEqual(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
                AssertWorkbookOpensViaExcelComWhenAvailable(
                    path,
                    "The general BIFF8 workbook failed to open in desktop Excel.");
                using (LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(path)) {
                    Assert.False(result.Workbook.HasRefreshAllMarker);
                }
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
                Assert.Equal("Value", reader.GetName(0));
                Assert.True(reader.Read());
                Assert.Equal(1, reader.GetInt32(0));
            } finally {
                TryDelete(path);
            }
        }

        [Fact]
        public void LegacyXls_GeneralWriter_IndexesSparseCoordinateBandsAndFormattingOnlyRows() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Sparse");
            sheet.CellValue(10, 1, "First");
            sheet.SetRowHeight(65, 21d);
            sheet.CellValue(100, 1, "Last");

            byte[] workbook = document.ToBytes(
                ExcelFileFormat.Xls,
                new ExcelSaveOptions { DisableFastPackageWriter = true });

            AssertBiffIndexMatchesDimensions(
                workbook,
                expectedDataFirstRow: 9,
                expectedDataRowAfterLast: 100,
                expectedDimensionFirstRow: 9,
                expectedDimensionRowAfterLast: 100);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
                workbook,
                new ExcelReadOptions { HasHeaderRow = false });
            Assert.True(reader.Read());
            Assert.Equal("First", reader.GetString(0));
            int rowCount = 1;
            while (reader.Read()) {
                rowCount++;
            }
            Assert.Equal(91, rowCount);
        }

        [Fact]
        public void LegacyXls_GeneralWriter_LeavesIndexDataBoundsEmptyForFormattingOnlySheet() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("FormattingOnly");
            sheet.SetColumnWidth(1, 18d);
            sheet.SetRowHeight(65, 21d);

            byte[] workbook = document.ToBytes(
                ExcelFileFormat.Xls,
                new ExcelSaveOptions { DisableFastPackageWriter = true });

            AssertBiffIndexMatchesDimensions(
                workbook,
                expectedDataFirstRow: 0,
                expectedDataRowAfterLast: 0,
                expectedDimensionFirstRow: 0,
                expectedDimensionRowAfterLast: 65);
        }

        private static void AssertSingleValue(string path, int expected) {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal("Value", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal(expected, reader.GetInt32(0));
            int rowCount = 1;
            while (reader.Read()) {
                rowCount++;
            }
            Assert.Equal(255, rowCount);
        }
    }
}
