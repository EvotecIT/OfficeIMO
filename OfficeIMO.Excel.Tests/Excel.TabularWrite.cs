using System.Data;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void WriteObjects_WritesPackageNativeTypedRows() {
            using var output = new MemoryStream();
            var rows = new[] {
                new TabularWriteRow(1, "Zażółć", new DateTime(2026, 7, 10), true),
                new TabularWriteRow(2, null, new DateTime(2026, 7, 11), false)
            };

            ExcelDataSetImportResult result = ExcelDocument.WriteObjects(
                output,
                rows,
                new ExcelTabularColumn<TabularWriteRow>[] {
                    ExcelTabularColumn<TabularWriteRow>.Create("Id", row => row.Id),
                    ExcelTabularColumn<TabularWriteRow>.Create("Name", row => row.Name),
                    ExcelTabularColumn<TabularWriteRow>.Create("Created", row => row.Created),
                    ExcelTabularColumn<TabularWriteRow>.Create("Active", row => row.Active)
                },
                new ExcelTabularWriteOptions { SheetName = "Typed Rows" });

            Assert.Equal("Typed Rows", result.SheetName);
            Assert.Equal("A1:D3", result.Range);
            Assert.Equal(2, result.RowCount);
            Assert.Equal(4, result.ColumnCount);

            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet
                .Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Zażółć", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
            Assert.Equal(string.Empty, GetSpreadsheetCellText(spreadsheet, cells["B3"]));
            Assert.Equal("1", cells["D2"].CellValue!.Text);
            Assert.NotNull(cells["C2"].StyleIndex);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteDataReader_WritesPackageAndLeavesReaderOpen() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add(1, "Alpha");
            table.Rows.Add(2, "Beta");
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(output, reader);

            Assert.False(reader.IsClosed);
            Assert.Equal("A1:B3", result.Range);
            Assert.Equal(2, result.RowCount);
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet
                .Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
            Assert.Equal("2", cells["A3"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteDataReader_HeaderlessEmptyReaderWritesValidEmptySheet() {
            var table = new DataTable("Empty");
            table.Columns.Add("Id", typeof(int));
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions { IncludeHeaders = false });

            Assert.Equal(string.Empty, result.Range);
            Assert.Equal(0, result.RowCount);
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            Assert.Empty(spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteObjects_CompactPackageOmitsDataReferencesAndRoundTrips() {
            using var output = new MemoryStream();
            var rows = new[] {
                new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true),
                new TabularWriteRow(2, "Beta", new DateTime(2026, 7, 11), false)
            };

            ExcelDocument.WriteObjects(
                output,
                rows,
                new (string Header, Func<TabularWriteRow, object?> Selector)[] {
                    ("Id", row => row.Id),
                    ("Name", row => row.Name),
                    ("Created", row => row.Created),
                    ("Active", row => row.Active)
                },
                new ExcelTabularWriteOptions { IncludeCellReferences = false, UseSharedStrings = false });

            byte[] package = output.ToArray();
            using (var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false)) {
                var savedRows = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Row>().ToArray();
                Assert.All(savedRows[0].Elements<Cell>(), cell => Assert.NotNull(cell.CellReference));
                Assert.All(savedRows.Skip(1).SelectMany(row => row.Elements<Cell>()), cell => Assert.Null(cell.CellReference));
                Assert.Null(spreadsheet.WorkbookPart.SharedStringTablePart);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }

            using var reader = ExcelDocumentReader.Open(new MemoryStream(package, writable: false));
            object?[,] values = reader.GetSheet("Data").ReadRange("A1:D3");
            Assert.Equal("Id", values[0, 0]);
            Assert.Equal("Alpha", values[1, 1]);
            Assert.Equal(2, Convert.ToInt32(values[2, 0], CultureInfo.InvariantCulture));
            Assert.Equal(false, values[2, 3]);
        }

        [Fact]
        public void WriteObjects_RejectsDuplicateHeaders() {
            using var output = new MemoryStream();
            var rows = new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) };

            Assert.Throws<ArgumentException>(() => ExcelDocument.WriteObjects(
                output,
                rows,
                new (string Header, Func<TabularWriteRow, object?> Selector)[] {
                    ("Name", row => row.Id),
                    ("name", row => row.Name)
                }));
        }

        [Fact]
        public void WriteRows_StreamsTypedValuesAndRoundTrips() {
            using var output = new MemoryStream();
            var rows = new[] {
                new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10, 8, 30, 0), true),
                new TabularWriteRow(2, "Beta", new DateTime(2026, 7, 11, 9, 45, 0), false)
            };

            ExcelDataSetImportResult result = ExcelDocument.WriteRows(
                output,
                rows,
                ["Id", "Name", "Created", "Active"],
                static (writer, row) => writer
                    .Write(row.Id)
                    .Write(row.Name)
                    .Write(row.Created)
                    .Write(row.Active),
                new ExcelTabularWriteOptions { IncludeCellReferences = false, UseSharedStrings = false });

            Assert.Equal("A1:D3", result.Range);
            using var reader = ExcelDocumentReader.Open(new MemoryStream(output.ToArray(), writable: false));
            object?[,] values = reader.GetSheet("Data").ReadRange("A1:D3");
            Assert.Equal("Beta", values[2, 1]);
            Assert.Equal(false, values[2, 3]);
        }

        [Fact]
        public void WriteRows_RejectsRowsWithTheWrongCellCount() {
            using var output = new MemoryStream();

            Assert.Throws<InvalidOperationException>(() => ExcelDocument.WriteRows(
                output,
                new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) },
                ["Id", "Name"],
                static (writer, row) => writer.Write(row.Id),
                new ExcelTabularWriteOptions { UseSharedStrings = false }));
        }

        [Fact]
        public void WriteRows_RejectsRowsWithTooManyCells() {
            using var output = new MemoryStream();

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => ExcelDocument.WriteRows(
                output,
                new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) },
                ["Id"],
                static (writer, row) => writer.Write(row.Id).Write(row.Name),
                new ExcelTabularWriteOptions { UseSharedStrings = false }));

            Assert.Contains("more than 1 cells", exception.Message, StringComparison.Ordinal);
        }

        private sealed record TabularWriteRow(int Id, string? Name, DateTime Created, bool Active);
    }
}
