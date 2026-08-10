using OfficeIMO.Excel;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Xlsb_DirectTabularSave_WritesSupportedScalarValuesThroughPublicApi() {
            using ExcelDocument document = ExcelDocument.Create();
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

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
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
            while (reader.Read()) rowCount++;
            Assert.Equal(255, rowCount);
            Assert.False(reader.NextResult());
        }

        [Fact]
        public void Xlsb_DirectTabularSave_WritesRowSpansAcrossMultipleColumnSegments() {
            const int columnCount = 1_025;
            var table = new DataTable("Wide");
            var values = new object[columnCount];
            for (int column = 0; column < columnCount; column++) {
                table.Columns.Add("C" + (column + 1), typeof(int));
                values[column] = column + 1;
            }
            table.Rows.Add(values);
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            using ExcelDocument document = ExcelDocument.Create();
            document.InsertDataSet(dataSet, createTables: false, includeAutoFilter: false);
            Assert.True(document.HasDeferredDirectDataSetImport);

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            Assert.True(document.HasDeferredDirectDataSetImport);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.Equal(columnCount, reader.FieldCount);
            Assert.Equal("C1", reader.GetName(0));
            Assert.Equal("C1025", reader.GetName(columnCount - 1));
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.Equal(columnCount, reader.GetInt32(columnCount - 1));
            Assert.False(reader.Read());
        }

        [Fact]
        public void Xlsb_DirectTabularSave_FallsBackForDateValuesWithoutChangingSemantics() {
            DateTime expected = new DateTime(2026, 8, 10, 14, 30, 0, DateTimeKind.Unspecified);
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "When");
            sheet.CellValue(2, 1, expected);

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.NotEqual(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.True(reader.Read());
            Assert.Equal(expected, reader.GetDateTime(0));
            Assert.False(reader.Read());
        }

        [Fact]
        public async Task Xlsb_DirectTabularSave_CoversSyncAndAsyncStreamsAndMutableFileState() {
            static ExcelDocument CreateDocument() {
                ExcelDocument document = ExcelDocument.Create();
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                for (int row = 2; row <= 256; row++) sheet.CellValue(row, 1, row - 1);
                return document;
            }

            using (ExcelDocument syncDocument = CreateDocument()) {
                using var syncDestination = new MemoryStream();
                syncDocument.Save(syncDestination, ExcelFileFormat.Xlsb);
                Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, syncDocument.LastSaveDiagnostics.Writer);
                AssertSingleValue(syncDestination.ToArray(), 1);
            }

            using (ExcelDocument asyncDocument = CreateDocument()) {
                using var asyncDestination = new MemoryStream();
                await asyncDocument.SaveAsync(asyncDestination, ExcelFileFormat.Xlsb);
                Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, asyncDocument.LastSaveDiagnostics.Writer);
                AssertSingleValue(asyncDestination.ToArray(), 1);
            }

            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
            try {
                using ExcelDocument fileDocument = CreateDocument();
                fileDocument.Save(path);
                Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, fileDocument.LastSaveDiagnostics.Writer);
                AssertWorkbookOpensViaExcelComWhenAvailable(
                    path,
                    "The directly generated XLSB workbook failed to open in desktop Excel.");
                fileDocument.Sheets[0].CellValue(2, 1, 2);
                fileDocument.Save();
                AssertSingleValue(File.ReadAllBytes(path), 2);
            } finally {
                TryDelete(path);
            }
        }

        [Fact]
        public async Task Xlsb_DirectTabularSave_HonorsPreCanceledAsyncStreamWrite() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            sheet.CellValue(2, 1, 1);
            using var destination = new MemoryStream();
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                document.SaveAsync(destination, ExcelFileFormat.Xlsb, cancellationToken: cancellation.Token));
            Assert.Equal(0, destination.Length);
        }

        [Fact]
        public async Task Xlsb_DirectTabularSave_RejectsOverlongHeadersAndValuesBeforeWriting() {
            static ExcelDocument CreateDocument(string columnName, string value) {
                var table = new DataTable("Data");
                table.Columns.Add(columnName, typeof(string));
                table.Rows.Add(value);
                var dataSet = new DataSet();
                dataSet.Tables.Add(table);
                ExcelDocument document = ExcelDocument.Create();
                document.InsertDataSet(dataSet, createTables: false, includeAutoFilter: false);
                return document;
            }

            byte[] sentinel = Enumerable.Range(0, 64).Select(index => (byte)index).ToArray();
            using (ExcelDocument longHeader = CreateDocument(new string('H', 32_768), "Value")) {
                using var destination = new MemoryStream();
                destination.Write(sentinel, 0, sentinel.Length);
                ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                    longHeader.Save(destination, ExcelFileFormat.Xlsb));
                Assert.Contains("32,767", exception.Message, StringComparison.Ordinal);
                Assert.Equal(sentinel, destination.ToArray());
                Assert.Equal(sentinel.Length, destination.Position);
            }

            using (ExcelDocument longValue = CreateDocument("Value", new string('V', 32_768))) {
                using var destination = new MemoryStream();
                destination.Write(sentinel, 0, sentinel.Length);
                ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() =>
                    longValue.SaveAsync(destination, ExcelFileFormat.Xlsb));
                Assert.Contains("32,767", exception.Message, StringComparison.Ordinal);
                Assert.Equal(sentinel, destination.ToArray());
                Assert.Equal(sentinel.Length, destination.Position);
            }
        }

        [Fact]
        public void Xlsb_DirectTabularSave_StagesUnsupportedValueFallbackBeforeDestinationWrite() {
            Guid expected = Guid.NewGuid();
            var table = new DataTable("Data");
            table.Columns.Add("Id", typeof(Guid));
            table.Rows.Add(expected);
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            using ExcelDocument document = ExcelDocument.Create();
            document.InsertDataSet(dataSet, createTables: false, includeAutoFilter: false);
            using var destination = new MemoryStream();

            document.Save(destination, ExcelFileFormat.Xlsb);

            Assert.NotEqual(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(destination.ToArray());
            Assert.True(reader.Read());
            Assert.Equal(expected.ToString(), reader.GetString(0));
            Assert.False(reader.Read());
        }

        private static void AssertSingleValue(byte[] workbook, int expected) {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.Equal("Value", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal(expected, reader.GetInt32(0));
            int rowCount = 1;
            while (reader.Read()) rowCount++;
            Assert.Equal(255, rowCount);
        }
    }
}
