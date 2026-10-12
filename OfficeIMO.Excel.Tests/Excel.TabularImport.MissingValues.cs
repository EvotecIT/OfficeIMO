using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System.Data;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(ExcelFileFormat.Xlsx, false)]
        [InlineData(ExcelFileFormat.Xlsx, true)]
        [InlineData(ExcelFileFormat.Xls, false)]
        [InlineData(ExcelFileFormat.Xls, true)]
        [InlineData(ExcelFileFormat.Xlsb, false)]
        [InlineData(ExcelFileFormat.Xlsb, true)]
        public void TabularImport_DataSet_PreservesMissingValuesAndEmptyTextAcrossSavePaths(ExcelFileFormat format, bool standardWriter) {
            DateTime date = new DateTime(2026, 10, 8, 6, 0, 0);
            var table = new DataTable("Data");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Missing", typeof(string));
            table.Columns.Add("Empty", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));
            table.Rows.Add(1, DBNull.Value, string.Empty, date);
            table.Rows.Add(2, "Alpha", DBNull.Value, DBNull.Value);
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            using ExcelDocument document = ExcelDocument.Create();
            document.InsertDataSet(dataSet, createTables: false, includeAutoFilter: false);

            byte[] package = document.ToBytes(format, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package);
            Assert.Equal(new[] { "Id", "Missing", "Empty", "Date" }, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.True(reader.IsDBNull(1));
            Assert.False(reader.IsDBNull(2));
            Assert.Equal(string.Empty, reader.GetString(2));
            Assert.Equal(date, reader.GetDateTime(3));
            Assert.True(reader.Read());
            Assert.Equal(2, reader.GetInt32(0));
            Assert.Equal("Alpha", reader.GetString(1));
            Assert.True(reader.IsDBNull(2));
            Assert.True(reader.IsDBNull(3));
            Assert.False(reader.Read());
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Contains("yy", loaded.Sheets[0].GetCellStyle(2, 4).NumberFormatCode, StringComparison.OrdinalIgnoreCase);
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsx, ExcelExecutionMode.Sequential, false)]
        [InlineData(ExcelFileFormat.Xlsx, ExcelExecutionMode.Sequential, true)]
        [InlineData(ExcelFileFormat.Xlsx, ExcelExecutionMode.Parallel, true)]
        [InlineData(ExcelFileFormat.Xlsb, ExcelExecutionMode.Sequential, false)]
        [InlineData(ExcelFileFormat.Xlsb, ExcelExecutionMode.Sequential, true)]
        [InlineData(ExcelFileFormat.Xlsb, ExcelExecutionMode.Parallel, true)]
        public void TabularImport_DataTable_AppendAndOverwritePreserveMissingValues(ExcelFileFormat format, ExcelExecutionMode mode, bool overwrite) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            // A second worksheet requires ordinary DOM insertion rather than a deferred source.
            document.AddWorksheet("Other");
            if (overwrite) {
                sheet.CellValue(2, 1, "Previous");
                sheet.CellFormula(2, 2, "123");
                sheet.FormatCell(2, 2, "0.0000");
                Cell inlineCell = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
                inlineCell.CellValue = null;
                inlineCell.DataType = CellValues.InlineString;
                inlineCell.InlineString = new InlineString(new Text("Previous"));
            }
            var table = new DataTable("Data");
            table.Columns.Add("Missing", typeof(string));
            table.Columns.Add("Number", typeof(int));
            table.Columns.Add("Empty", typeof(string));
            table.Rows.Add(DBNull.Value, DBNull.Value, string.Empty);
            table.Rows.Add("Alpha", 2, "Beta");
            using var cancellation = new CancellationTokenSource();

            sheet.InsertDataTable(table, mode: mode, ct: cancellation.Token);
            byte[] package = document.ToBytes(format);

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { SheetName = "Data" });
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.IsDBNull(1));
            Assert.False(reader.IsDBNull(2));
            Assert.Equal(string.Empty, reader.GetString(2));
            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetString(0));
            Assert.Equal(2, reader.GetInt32(1));
            Assert.Equal("Beta", reader.GetString(2));
            Assert.False(reader.Read());
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Null(loaded.Sheets[0].CellAt(2, 2).GetValue().Formula);
            if (overwrite) Assert.Equal("0.0000", loaded.Sheets[0].GetCellStyle(2, 2).NumberFormatCode);
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsx, false)]
        [InlineData(ExcelFileFormat.Xlsx, true)]
        [InlineData(ExcelFileFormat.Xlsb, false)]
        [InlineData(ExcelFileFormat.Xlsb, true)]
        public void TabularImport_ProjectedObjects_PreserveNullAndDbNullAsMissing(ExcelFileFormat format, bool standardWriter) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.InsertObjects(new[] { 1, 2, 3 },
                ("Id", (Func<int, object?>)(id => id)),
                ("Text", id => id == 1 ? null : id == 2 ? DBNull.Value : string.Empty));

            byte[] package = document.ToBytes(format, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });

            if (format == ExcelFileFormat.Xlsx && standardWriter) AssertMissingTabularDefaultStyles(package, "B2", "B3");
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package);
            for (int id = 1; id <= 3; id++) {
                Assert.True(reader.Read());
                Assert.Equal(id, reader.GetInt32(0));
                Assert.Equal(id < 3, reader.IsDBNull(1));
                if (id == 3) Assert.Equal(string.Empty, reader.GetString(1));
            }
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData(ExcelExecutionMode.Sequential)]
        [InlineData(ExcelExecutionMode.Parallel)]
        public void TabularImport_MissingOverwriteRegistersDefaultStyleForStandardXlsx(ExcelExecutionMode mode) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Previous");
            DataTable table = new DataTable();
            table.Columns.Add("Missing", typeof(string));
            table.Columns.Add("AlsoMissing", typeof(string));
            table.Columns.Add("Empty", typeof(string));
            table.Rows.Add(DBNull.Value, DBNull.Value, string.Empty);

            sheet.InsertDataTable(table, includeHeaders: false, mode: mode);
            byte[] package = document.ToBytes(ExcelFileFormat.Xlsx, new ExcelSaveOptions { DisableFastPackageWriter = true });

            AssertMissingTabularDefaultStyles(package, "A1", "B1");
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { HasHeaderRow = false });
            Assert.Equal(3, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.IsDBNull(1));
            Assert.False(reader.IsDBNull(2));
            Assert.Equal(string.Empty, reader.GetString(2));
            Assert.False(reader.Read());
        }

        private static void AssertMissingTabularDefaultStyles(byte[] package, params string[] references) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
            WorkbookPart workbook = spreadsheet.WorkbookPart!;
            CellFormats formats = workbook.WorkbookStylesPart!.Stylesheet!.CellFormats!;
            Assert.NotEmpty(formats.Elements<CellFormat>());
            Dictionary<string, Cell> cells = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            foreach (string reference in references) {
                Cell missing = cells[reference];
                Assert.Equal(0U, missing.StyleIndex!.Value);
                Assert.Null(missing.CellValue);
                Assert.Null(missing.DataType);
            }
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsx, false)]
        [InlineData(ExcelFileFormat.Xlsx, true)]
        [InlineData(ExcelFileFormat.Xlsb, false)]
        [InlineData(ExcelFileFormat.Xlsb, true)]
        public void TabularImport_ProjectedFallback_PreservesMissingValuesWhenAppendingOrOverwriting(ExcelFileFormat format, bool overwrite) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            if (overwrite) sheet.CellFormula(2, 2, "123");
            sheet.InsertObjects(Enumerable.Range(1, 20),
                ("Id", (Func<int, object?>)(id => id)),
                ("Text", id => id == 1 ? null : id == 2 ? DBNull.Value : string.Empty));

            byte[] package = document.ToBytes(format);

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { SheetName = "Data" });
            for (int id = 1; id <= 20; id++) {
                Assert.True(reader.Read());
                Assert.Equal(id, reader.GetInt32(0));
                Assert.Equal(id <= 2, reader.IsDBNull(1));
                if (id > 2) Assert.Equal(string.Empty, reader.GetString(1));
            }
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsx)]
        [InlineData(ExcelFileFormat.Xls)]
        [InlineData(ExcelFileFormat.Xlsb)]
        public void TabularImport_DataReader_PreservesMissingAndEmptyText(ExcelFileFormat format) {
            var table = new DataTable();
            table.Columns.Add("Missing", typeof(string));
            table.Columns.Add("Empty", typeof(string));
            table.Rows.Add(DBNull.Value, string.Empty);
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            using DataTableReader source = table.CreateDataReader();
            sheet.InsertDataReader(source, createTable: false, includeAutoFilter: false);

            byte[] package = document.ToBytes(format);

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { SheetName = "Data" });
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.False(reader.IsDBNull(1));
            Assert.Equal(string.Empty, reader.GetString(1));
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsx, false)]
        [InlineData(ExcelFileFormat.Xlsx, true)]
        [InlineData(ExcelFileFormat.Xls, false)]
        [InlineData(ExcelFileFormat.Xls, true)]
        [InlineData(ExcelFileFormat.Xlsb, false)]
        [InlineData(ExcelFileFormat.Xlsb, true)]
        public void TabularImport_HeaderlessMissingCells_PreserveDeclaredShape(ExcelFileFormat format, bool standardWriter) {
            var table = new DataTable("Data");
            table.Columns.Add("A", typeof(string));
            table.Columns.Add("B", typeof(string));
            for (int row = 0; row < 4; row++) table.Rows.Add(DBNull.Value, DBNull.Value);
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            using ExcelDocument document = ExcelDocument.Create();
            document.InsertDataSet(dataSet, createTables: false, includeHeaders: false, includeAutoFilter: false);

            byte[] package = document.ToBytes(format, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { HasHeaderRow = false });
            Assert.Equal(2, reader.FieldCount);
            for (int row = 0; row < 4; row++) {
                Assert.True(reader.Read());
                Assert.True(reader.IsDBNull(0));
                Assert.True(reader.IsDBNull(1));
            }
            Assert.False(reader.Read());
        }

        [Fact]
        public void TabularImport_ExplicitCellNullCompatibility_SurvivesDeferredMaterialization() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            for (int row = 1; row <= 140; row++) {
                sheet.CellValue(row, 1, row == 2 ? DBNull.Value : row == 3 ? null : (object)"Value");
            }

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsx, new ExcelSaveOptions { DisableFastPackageWriter = true });

            using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            foreach (string reference in new[] { "A2", "A3" }) {
                Assert.Equal(CellValues.String, cells[reference].DataType!.Value);
                Assert.Equal(string.Empty, cells[reference].CellValue!.Text);
            }
        }

        [Fact]
        public void TabularImport_MissingAndEmptyText_OpenCorrectlyInDesktopExcelWhenAvailable() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
            try {
                var table = new DataTable("Data");
                table.Columns.Add("Text", typeof(string));
                table.Rows.Add(DBNull.Value);
                table.Rows.Add(string.Empty);
                using ExcelDocument document = ExcelDocument.Create();
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(table);
                sheet.CellFormula(1, 2, "ISBLANK(A2)");
                sheet.CellFormula(1, 3, "ISTEXT(A3)");
                document.Save(path, new ExcelSaveOptions { XlsbUseSharedStrings = true, DisableFastPackageWriter = true });

                AssertWorkbookOpensViaExcelComWhenAvailable(path,
                    "Desktop Excel must distinguish a missing imported cell from explicit empty text.",
                    new Dictionary<string, string> { ["B1"] = "TRUE", ["C1"] = "TRUE" });
            } finally {
                TryDelete(path);
            }
        }
    }
}
