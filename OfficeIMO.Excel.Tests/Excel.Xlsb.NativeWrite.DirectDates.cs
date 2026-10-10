using OfficeIMO.Excel;
using System.Data;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(ExcelDateSystem.NineteenHundred, false, false)]
        [InlineData(ExcelDateSystem.NineteenFour, false, true)]
        [InlineData(ExcelDateSystem.NineteenHundred, true, true)]
        [InlineData(ExcelDateSystem.NineteenFour, true, false)]
        public void Xlsb_DirectDates_PreserveInsertionValuesAndFormats(
            ExcelDateSystem dateSystem, bool dataTable, bool sharedStrings) {
            var rows = new[] {
                new DirectDateRecord { Name = "early", When = new DateTime(1900, 2, 28, 6, 30, 0), Value = 42.5 },
                new DirectDateRecord { Name = "epoch", When = new DateTime(1904, 1, 1), Value = -1.25 },
                new DirectDateRecord { Name = "modern", When = new DateTime(2026, 10, 9, 14, 30, 15, 123), Value = 9.5 },
            };
            using ExcelDocument direct = CreateDirectDateDocument(rows, dateSystem, dataTable);
            var diagnostics = new List<string>();
            direct.Execution.OnInfo = diagnostics.Add;
            byte[] bytes = direct.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings });

            Assert.True(direct.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.NativeBinaryDirectPackage,
                string.Join("; ", diagnostics));
            Assert.True(direct.HasDeferredDirectDataSetImport);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(bytes);
            Assert.Equal(new[] { "Name", "When", "Value" }, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
            foreach (DirectDateRecord row in rows) {
                Assert.True(reader.Read());
                Assert.Equal(row.Name, reader.GetString(0));
                Assert.Equal(row.When, reader.GetDateTime(1));
                Assert.Equal(row.Value, reader.GetDouble(2));
            }
            Assert.False(reader.Read());
            Assert.False(reader.NextResult());

            using ExcelDocument standard = CreateDirectDateDocument(rows, dateSystem, dataTable);
            byte[] standardBytes = standard.ToBytes(ExcelFileFormat.Xlsb,
                new ExcelSaveOptions { DisableFastPackageWriter = true, XlsbUseSharedStrings = sharedStrings });
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(bytes, writable: false));
            using ExcelDocument loadedStandard = ExcelDocument.Load(new MemoryStream(standardBytes, writable: false));
            Assert.Equal(dateSystem, loaded.DateSystem);
            for (int row = 2; row <= rows.Length + 1; row++) {
                Assert.Equal(loadedStandard.Sheets[0].GetCellStyle(row, 2).NumberFormatCode,
                    loaded.Sheets[0].GetCellStyle(row, 2).NumberFormatCode);
                Assert.Equal(AssertCellValue(loadedStandard.Sheets[0], row, 2).DateTimeValue,
                    AssertCellValue(loaded.Sheets[0], row, 2).DateTimeValue);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_DirectDates_PreserveMixedMissingValuesAndUnsupportedFallback(bool unsupportedValue) {
            DateTime date = new DateTime(2026, 10, 9, 6, 0, 0);
            Guid id = new Guid("3cd8c138-9492-4934-9529-a1f89f586f32");
            var table = new DataTable("Data");
            table.Columns.Add("Value", typeof(object));
            table.Rows.Add(date);
            table.Rows.Add(DBNull.Value);
            table.Rows.Add("");
            table.Rows.Add(42D);
            if (unsupportedValue) table.Rows.Add(id);
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            using ExcelDocument document = ExcelDocument.Create();
            document.InsertDataSet(dataSet, createTables: false, includeAutoFilter: false);
            using var destination = new MemoryStream();
            document.Save(destination, ExcelFileFormat.Xlsb);
            if (!unsupportedValue) {
                Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            }

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(destination.ToArray());
            Assert.True(reader.Read());
            Assert.Equal(date, reader.GetDateTime(0));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.Read());
            Assert.False(reader.IsDBNull(0));
            Assert.Equal("", reader.GetString(0));
            Assert.True(reader.Read());
            Assert.Equal(42D, reader.GetDouble(0));
            if (unsupportedValue) {
                Assert.True(reader.Read());
                Assert.Equal(id.ToString(), reader.GetString(0));
            }
            Assert.False(reader.Read());
        }

        private static ExcelDocument CreateDirectDateDocument(
            DirectDateRecord[] rows, ExcelDateSystem dateSystem, bool dataTable) {
            ExcelDocument document = ExcelDocument.Create();
            document.DateSystem = dateSystem;
            if (dataTable) {
                var table = new DataTable("Data");
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("When", typeof(DateTime));
                table.Columns.Add("Value", typeof(double));
                foreach (DirectDateRecord row in rows) table.Rows.Add(row.Name, row.When, row.Value);
                var dataSet = new DataSet();
                dataSet.Tables.Add(table);
                document.InsertDataSet(dataSet, createTables: false, includeAutoFilter: false);
            } else {
                document.AddWorksheet("Data").InsertObjects(rows,
                    ("Name", static row => row.Name), ("When", static row => row.When), ("Value", static row => row.Value));
            }
            return document;
        }

        private sealed class DirectDateRecord {
            public string Name { get; set; } = "";
            public DateTime When { get; set; }
            public double Value { get; set; }
        }
    }
}
