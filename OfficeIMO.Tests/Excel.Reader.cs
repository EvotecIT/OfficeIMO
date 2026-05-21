using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_EnumerateCells_ReturnsCorrectCoordinates() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderEnumerateCellsCoordinates.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "B2");
                    sheet.CellValue(3, 4, "D3");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var cells = reader.GetSheet("Data").EnumerateCells().ToList();

                Assert.Contains(cells, c => c.Row == 2 && c.Column == 2 && Equals(c.Value, "B2"));
                Assert.Contains(cells, c => c.Row == 3 && c.Column == 4 && Equals(c.Value, "D3"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_OpenPath_DetachesFromSourceFile() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderOpenPathDetached.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                File.Delete(filePath);
                Assert.False(File.Exists(filePath));

                object?[,] values = reader.GetSheet("Data").ReadRange("A1:A1");
                Assert.Equal("Value", values[0, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_OpenStream_CopiesSeekableStreamAndLeavesSourceOpen() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderOpenStreamCopiesSeekable.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    document.Save();
                }

                byte[] bytes = File.ReadAllBytes(filePath);
                using var stream = new MemoryStream(bytes, 0, bytes.Length, writable: true, publiclyVisible: true);
                stream.Position = stream.Length;

                using (var reader = ExcelDocumentReader.Open(stream)) {
                    Array.Clear(stream.GetBuffer(), 0, Math.Min(16, stream.GetBuffer().Length));
                    object?[,] values = reader.GetSheet("Data").ReadRange("A1:A1");

                    Assert.Equal("Value", values[0, 0]);
                }

                Assert.True(stream.CanRead);
                stream.Position = 0;
                Assert.Equal(0, stream.Position);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRange_UsesConfiguredCultureBeforeInvariantNumericFallback() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderCultureNumericFallback.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 1d);
                    sheet.CellValue(1, 2, 2d);
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
                    cells["A1"].DataType = CellValues.Number;
                    cells["A1"].CellValue = new CellValue("1,23");
                    cells["B1"].DataType = CellValues.Number;
                    cells["B1"].CellValue = new CellValue("123.45");
                    spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
                }

                var options = new ExcelReadOptions {
                    Culture = CultureInfo.GetCultureInfo("pl-PL")
                };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:B1");

                Assert.Equal(1.23d, Assert.IsType<double>(values[0, 0]), precision: 2);
                Assert.Equal(123.45d, Assert.IsType<double>(values[0, 1]), precision: 2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_EnumerateRange_FiltersUsingActualColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderEnumerateRangeColumns.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "Inside");
                    sheet.CellValue(2, 4, "Outside");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var cells = reader.GetSheet("Data").EnumerateRange("B1:C3").ToList();

                var onlyCell = Assert.Single(cells);
                Assert.Equal(2, onlyCell.Row);
                Assert.Equal(2, onlyCell.Column);
                Assert.Equal("Inside", onlyCell.Value);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_EnumerateRange_DoesNotStopAtOutOfOrderRowsBeyondRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderEnumerateRangeOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 1, "InRange");
                    sheet.CellValue(5, 1, "Outside");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!;
                    var row2 = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 2U);
                    row2.Remove();
                    sheetData.Append(row2);
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var cells = reader.GetSheet("Data").EnumerateRange("A1:A2").ToList();

                Assert.Contains(cells, c => c.Row == 1 && c.Column == 1 && Equals(c.Value, "Header"));
                Assert.Contains(cells, c => c.Row == 2 && c.Column == 1 && Equals(c.Value, "InRange"));
                Assert.DoesNotContain(cells, c => c.Row == 5);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_RowReaders_DoNotStopAtOutOfOrderRowsBeyondRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowReadersOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 1, "InRange");
                    sheet.CellValue(5, 1, "Outside");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 2U);

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                var rows = sheetReader.ReadRows("A1:A2").ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal("Header", rows[0]![0]);
                Assert.Equal("InRange", rows[1]![0]);

                var column = sheetReader.ReadColumn("A1:A2").ToList();
                Assert.Equal(new object?[] { "Header", "InRange" }, column);

                var sequentialChunks = sheetReader
                    .ReadRangeStream("A1:A2", chunkRows: 1, mode: ExecutionMode.Sequential)
                    .ToList();
                Assert.Equal(new[] { 1, 2 }, sequentialChunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal("InRange", sequentialChunks[1].Rows[0][0]);

                var parallelChunks = sheetReader
                    .ReadRangeStream("A1:A2", chunkRows: 1, mode: ExecutionMode.Parallel)
                    .ToList();
                Assert.Equal(new[] { 1, 2 }, parallelChunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal("InRange", parallelChunks[1].Rows[0][0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_FastReaders_DoNotStopAtOutOfOrderRowsBeyondRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFastOutOfOrderRowsBeyondRange.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 1, "InRange");
                    sheet.CellValue(5, 1, "Outside");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 2U);

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                object?[,] range = sheetReader.ReadRange("A1:A2");
                Assert.Equal("Header", range[0, 0]);
                Assert.Equal("InRange", range[1, 0]);

                var column = sheetReader.ReadColumn("A1:A2").ToArray();
                Assert.Equal(new object?[] { "Header", "InRange" }, column);

                var table = sheetReader.ReadRangeAsDataTable("A1:A2", headersInFirstRow: false);
                Assert.Equal("Header", table.Rows[0][0]);
                Assert.Equal("InRange", table.Rows[1][0]);

                var singleChunk = Assert.Single(sheetReader.ReadRangeStream("A1:A2", chunkRows: 2));
                Assert.Equal("Header", singleChunk.Rows[0][0]);
                Assert.Equal("InRange", singleChunk.Rows[1][0]);

                var bufferedChunks = sheetReader.ReadRangeStream("A1:A2", chunkRows: 1).ToList();
                Assert.Equal(new[] { 1, 2 }, bufferedChunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal("InRange", bufferedChunks[1].Rows[0][0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_RowReaders_HandleOutOfOrderCellsWithinRow() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowReadersOutOfOrderCells.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "A");
                    sheet.CellValue(1, 2, "B");
                    sheet.CellValue(1, 3, "C");
                    sheet.CellValue(1, 4, "Outside");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var row = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Single(r => r.RowIndex?.Value == 1U);
                    var cells = row.Elements<Cell>().ToDictionary(c => c.CellReference!.Value!);
                    row.RemoveAllChildren<Cell>();
                    row.Append(cells["C1"]);
                    row.Append(cells["A1"]);
                    row.Append(cells["D1"]);
                    row.Append(cells["B1"]);
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                object?[,] range = sheetReader.ReadRange("A1:C1");
                Assert.Equal("A", range[0, 0]);
                Assert.Equal("B", range[0, 1]);
                Assert.Equal("C", range[0, 2]);

                object?[] rowValues = Assert.Single(sheetReader.ReadRows("A1:C1"));
                Assert.Equal(new object?[] { "A", "B", "C" }, rowValues);

                var streamChunk = Assert.Single(sheetReader.ReadRangeStream("A1:C1", chunkRows: 1));
                Assert.Equal(new object?[] { "A", "B", "C" }, streamChunk.Rows[0]);

                var table = sheetReader.ReadRangeAsDataTable("A1:C1", headersInFirstRow: false);
                Assert.Equal("A", table.Rows[0][0]);
                Assert.Equal("B", table.Rows[0][1]);
                Assert.Equal("C", table.Rows[0][2]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_DataTable_NoHeadersNoInference_UsesGeneratedObjectColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableNoHeadersNoInference.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Count");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 42);
                    document.Save();
                }

                var options = new ExcelReadOptions { InferDataTableColumnTypes = false };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:B2", headersInFirstRow: false, mode: ExecutionMode.Sequential);

                Assert.Equal(new[] { "Column1", "Column2" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
                Assert.All(table.Columns.Cast<DataColumn>(), column => Assert.Equal(typeof(object), column.DataType));
                Assert.Equal("Name", table.Rows[0][0]);
                Assert.Equal("Count", table.Rows[0][1]);
                Assert.Equal("Alpha", table.Rows[1][0]);
                Assert.Equal(42D, table.Rows[1][1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_DataTable_HeadersNoInference_UsesHeaderObjectColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableHeadersNoInference.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Count");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 42);
                    sheet.CellValue(3, 1, "Beta");
                    sheet.CellValue(3, 2, 7);
                    document.Save();
                }

                var options = new ExcelReadOptions { InferDataTableColumnTypes = false };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:B3", headersInFirstRow: true, mode: ExecutionMode.Sequential);

                Assert.Equal(new[] { "Name", "Count" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
                Assert.All(table.Columns.Cast<DataColumn>(), column => Assert.Equal(typeof(object), column.DataType));
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal("Alpha", table.Rows[0][0]);
                Assert.Equal(42D, table.Rows[0][1]);
                Assert.Equal("Beta", table.Rows[1][0]);
                Assert.Equal(7D, table.Rows[1][1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_DataTable_MixedTypeInference_ResolvesObjectColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableMixedTypeInference.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Key");
                    sheet.CellValue(1, 2, "Value");
                    sheet.CellValue(2, 1, 1);
                    sheet.CellValue(2, 2, "Open");
                    sheet.CellValue(3, 1, "Two");
                    sheet.CellValue(3, 2, 7);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:B3", headersInFirstRow: true, mode: ExecutionMode.Sequential);

                Assert.Equal(new[] { "Key", "Value" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
                Assert.All(table.Columns.Cast<DataColumn>(), column => Assert.Equal(typeof(object), column.DataType));
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal(1D, table.Rows[0][0]);
                Assert.Equal("Open", table.Rows[0][1]);
                Assert.Equal("Two", table.Rows[1][0]);
                Assert.Equal(7D, table.Rows[1][1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_RowReaders_HandleLargeSortedSparseRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowReadersLargeSortedSparseRanges.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 2, "OutsideRequestedColumn");
                    sheet.CellValue(100001, 1, "Tail");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                var column = sheetReader.ReadColumn("A1:A100001").ToArray();
                Assert.Equal(100001, column.Length);
                Assert.Equal("Header", column[0]);
                Assert.Null(column[1]);
                Assert.Equal("Tail", column[100000]);

                var rows = sheetReader.ReadRows("A1:A100001").ToArray();
                Assert.Equal(100001, rows.Length);
                Assert.Equal("Header", rows[0]![0]);
                Assert.Null(rows[1]);
                Assert.Equal("Tail", rows[100000]![0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRows_LargeSparseOutOfOrderRowsRemainOrdered() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowsLargeSparseOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 2, "OutsideRequestedColumn");
                    sheet.CellValue(100001, 1, "Tail");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 1U);

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadRows("A1:A100001").ToArray();

                Assert.Equal(100001, rows.Length);
                Assert.Equal("Header", rows[0]![0]);
                Assert.Null(rows[1]);
                Assert.Equal("Tail", rows[100000]![0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadObjects_SequentialOutOfOrderRowsRemainOrdered() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderObjectsOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Count");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 42);
                    sheet.CellValue(4, 1, "Omega");
                    sheet.CellValue(4, 2, 99);
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 1U);

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadObjects("A1:B4", ExecutionMode.Sequential).ToList();

                Assert.Equal(3, rows.Count);
                Assert.Equal("Alpha", rows[0]["Name"]);
                Assert.Equal(42D, rows[0]["Count"]);
                Assert.Null(rows[1]["Name"]);
                Assert.Null(rows[1]["Count"]);
                Assert.Equal("Omega", rows[2]["Name"]);
                Assert.Equal(99D, rows[2]["Count"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadObjects_SequentialHonorsCellValueConverter() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderObjectsCellValueConverter.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Count");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 42);
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    CellValueConverter = context => context.RawText == "42" ? new ExcelCellValue("forty-two") : ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:B2", ExecutionMode.Sequential));

                Assert.Equal("Alpha", row["Name"]);
                Assert.Equal("forty-two", row["Count"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRows_HonorsCellValueConverter() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowsCellValueConverter.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Count");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 42);
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    CellValueConverter = context => context.RawText == "42" ? new ExcelCellValue("forty-two") : ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadRows("A2:B2"));

                Assert.Equal("Alpha", row![0]);
                Assert.Equal("forty-two", row[1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadColumn_HonorsCellValueConverter() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderColumnCellValueConverter.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, 42);
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    CellValueConverter = context => context.RawText == "42" ? new ExcelCellValue("forty-two") : ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var column = reader.GetSheet("Data").ReadColumn("A1:A2").ToList();

                Assert.Equal("Name", column[0]);
                Assert.Equal("forty-two", column[1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadColumn_CellValueConverterFallbackUsesConfiguredCulture() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderColumnConverterCultureFallback.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 1d);
                    sheet.CellValue(2, 1, 2d);
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
                    cells["A1"].DataType = CellValues.Number;
                    cells["A1"].CellValue = new CellValue("1,23");
                    cells["A2"].DataType = CellValues.Number;
                    cells["A2"].CellValue = new CellValue("123.45");
                    spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
                }

                var options = new ExcelReadOptions {
                    Culture = CultureInfo.GetCultureInfo("pl-PL"),
                    CellValueConverter = static _ => ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var column = reader.GetSheet("Data").ReadColumn("A1:A2").ToList();

                Assert.Equal(1.23d, Assert.IsType<double>(column[0]), precision: 2);
                Assert.Equal(123.45d, Assert.IsType<double>(column[1]), precision: 2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRows_CellValueConverterFallbackUsesConfiguredCulture() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowsConverterCultureFallback.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 1d);
                    sheet.CellValue(2, 1, 2d);
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
                    cells["A1"].DataType = CellValues.Number;
                    cells["A1"].CellValue = new CellValue("1,23");
                    cells["A2"].DataType = CellValues.Number;
                    cells["A2"].CellValue = new CellValue("123.45");
                    spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
                }

                var options = new ExcelReadOptions {
                    Culture = CultureInfo.GetCultureInfo("pl-PL"),
                    CellValueConverter = static _ => ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var rows = reader.GetSheet("Data").ReadRows("A1:A2").ToList();

                Assert.Equal(1.23d, Assert.IsType<double>(rows[0]![0]), precision: 2);
                Assert.Equal(123.45d, Assert.IsType<double>(rows[1]![0]), precision: 2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_HonorsCellValueConverter() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamCellValueConverter.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Count");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 42);
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    CellValueConverter = context => context.RawText == "42" ? new ExcelCellValue("forty-two") : ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var chunks = reader.GetSheet("Data").ReadRangeStream("A1:B2", chunkRows: 1, mode: ExecutionMode.Sequential).ToList();

                Assert.Equal(2, chunks.Count);
                Assert.Equal("Alpha", chunks[1].Rows[0][0]);
                Assert.Equal("forty-two", chunks[1].Rows[0][1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_CellValueConverterFallbackUsesConfiguredCulture() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamConverterCultureFallback.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 1d);
                    sheet.CellValue(2, 1, 2d);
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
                    cells["A1"].DataType = CellValues.Number;
                    cells["A1"].CellValue = new CellValue("1,23");
                    cells["A2"].DataType = CellValues.Number;
                    cells["A2"].CellValue = new CellValue("123.45");
                    spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
                }

                var options = new ExcelReadOptions {
                    Culture = CultureInfo.GetCultureInfo("pl-PL"),
                    CellValueConverter = static _ => ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var chunk = Assert.Single(reader.GetSheet("Data").ReadRangeStream("A1:A2", chunkRows: 2, mode: ExecutionMode.Sequential));

                Assert.Equal(1.23d, Assert.IsType<double>(chunk.Rows[0][0]), precision: 2);
                Assert.Equal(123.45d, Assert.IsType<double>(chunk.Rows[1][0]), precision: 2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_SequentialHonorsCellValueConverter() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableCellValueConverter.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Count");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 42);
                    sheet.CellValue(3, 1, "Beta");
                    sheet.CellValue(3, 2, 7);
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    CellValueConverter = context => context.RawText == "42" ? new ExcelCellValue("forty-two") : ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:B3", mode: ExecutionMode.Sequential);

                Assert.Equal("Alpha", table.Rows[0]["Name"]);
                Assert.Equal("forty-two", table.Rows[0]["Count"]);
                Assert.Equal("Beta", table.Rows[1]["Name"]);
                Assert.Equal(7d, table.Rows[1]["Count"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_CellValueConverterFallbackUsesConfiguredCulture() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableConverterCultureFallback.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Amount");
                    sheet.CellValue(2, 1, 1d);
                    sheet.CellValue(3, 1, 2d);
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
                    cells["A2"].DataType = CellValues.Number;
                    cells["A2"].CellValue = new CellValue("1,23");
                    cells["A3"].DataType = CellValues.Number;
                    cells["A3"].CellValue = new CellValue("123.45");
                    spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
                }

                var options = new ExcelReadOptions {
                    Culture = CultureInfo.GetCultureInfo("pl-PL"),
                    CellValueConverter = static _ => ExcelCellValue.NotHandled
                };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:A3", mode: ExecutionMode.Sequential);

                Assert.Equal(1.23d, Assert.IsType<double>(table.Rows[0]["Amount"]), precision: 2);
                Assert.Equal(123.45d, Assert.IsType<double>(table.Rows[1]["Amount"]), precision: 2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static void MoveWorksheetRowToEnd(string filePath, uint rowIndex) {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!;
            var row = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == rowIndex);
            row.Remove();
            sheetData.Append(row);
            worksheetPart.Worksheet.Save();
        }

        [Fact]
        public void Reader_FormulaText_IsReturnedWhenCachedResultsAreDisabled() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFormulaText.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 2);
                    sheet.CellValue(2, 1, 3);
                    sheet.CellFormula(3, 1, "=SUM(A1:A2)");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { UseCachedFormulaResult = false });
                var cell = reader.GetSheet("data").EnumerateCells().Single(c => c.Row == 3 && c.Column == 1);

                Assert.Equal("SUM(A1:A2)", cell.Value);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_FormulaText_IsReturnedWhenCachedResultIsMissing() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFormulaWithoutCachedResult.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 2);
                    sheet.CellValue(2, 1, 3);
                    sheet.CellFormula(3, 1, "=SUM(A1:A2)");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var cell = reader.GetSheet("Data").EnumerateCells().Single(c => c.Row == 3 && c.Column == 1);

                Assert.Equal("SUM(A1:A2)", cell.Value);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeSequential_FormulaText_IsReturnedWhenCachedResultsAreDisabled() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderSequentialFormulaText.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 2);
                    sheet.CellValue(2, 1, 3);
                    sheet.CellFormula(3, 1, "=SUM(A1:A2)");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { UseCachedFormulaResult = false });
                object?[,] values = reader.GetSheet("Data").ReadRange("A3:A3", ExecutionMode.Sequential);

                Assert.Equal("SUM(A1:A2)", values[0, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAutomatic_FormulaText_IsReturnedWhenCachedResultsAreDisabled() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderAutomaticFormulaText.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 2);
                    sheet.CellValue(2, 1, 3);
                    sheet.CellFormula(3, 1, "=SUM(A1:A2)");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { UseCachedFormulaResult = false });
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:A3");

                Assert.Equal(2D, values[0, 0]);
                Assert.Equal(3D, values[1, 0]);
                Assert.Equal("SUM(A1:A2)", values[2, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAutomatic_FormulaText_SkipsCachedValueWhenDisabled() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFormulaTextWithCachedValue.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 2);
                    sheet.CellValue(2, 1, 3);
                    sheet.CellFormula(3, 1, "=SUM(A1:A2)");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                    var formulaCell = worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == "A3");
                    formulaCell.CellValue = new CellValue("5");
                    worksheet.Save();
                }

                using var cachedReader = ExcelDocumentReader.Open(filePath);
                object?[,] cachedValues = cachedReader.GetSheet("Data").ReadRange("A3:A3");
                Assert.Equal(5D, cachedValues[0, 0]);

                using var formulaReader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { UseCachedFormulaResult = false });
                object?[,] formulaValues = formulaReader.GetSheet("Data").ReadRange("A3:A3");
                Assert.Equal("SUM(A1:A2)", formulaValues[0, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeSequential_HonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderSequentialCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                using var cts = new CancellationTokenSource();
                cts.Cancel();

                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadRange("A1:A1", ExecutionMode.Sequential, cts.Token));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadColumn_HonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderColumnCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 42);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                using var cts = new CancellationTokenSource();
                cts.Cancel();

                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadColumn("A1:A1", cts.Token).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRows_HonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowsCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 42);
                    sheet.CellValue(1, 2, 43);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                using var cts = new CancellationTokenSource();
                cts.Cancel();

                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadRows("A1:B1", cts.Token).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadColumnAs_HonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderColumnAsCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 42);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                using var cts = new CancellationTokenSource();
                cts.Cancel();

                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadColumnAs<int>("A1:A1", ct: cts.Token).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRowsAs_HonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowsAsCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 42);
                    sheet.CellValue(1, 2, 43);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                using var cts = new CancellationTokenSource();
                cts.Cancel();

                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadRowsAs<int>("A1:B1", ct: cts.Token).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_HonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 42);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                using var cts = new CancellationTokenSource();
                cts.Cancel();

                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadRangeStream("A1:A1", ct: cts.Token).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_RejectsInvalidChunkRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamInvalidChunkRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    document.AddWorkSheet("Data").CellValue(1, 1, 42);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    reader.GetSheet("Data").ReadRangeStream("A1:A1", chunkRows: 0).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRange_ReportsOwnExecutionDecision() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeDecision.xlsx");
            var decisions = new List<(string Operation, int Items, ExecutionMode Mode)>();

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "A");
                    sheet.CellValue(1, 2, "B");
                    document.Save();
                }

                var options = new ExcelReadOptions();
                options.Execution.OperationThresholds["ReadRange"] = 1;
                options.Execution.OnDecision = (operation, items, mode) => decisions.Add((operation, items, mode));

                using var reader = ExcelDocumentReader.Open(filePath, options);
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:B1");

                Assert.Equal("A", values[0, 0]);
                Assert.Equal("B", values[0, 1]);
                var decision = Assert.Single(decisions);
                Assert.Equal("ReadRange", decision.Operation);
                Assert.Equal(2, decision.Items);
                Assert.Equal(ExecutionMode.Parallel, decision.Mode);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_ReportsOwnExecutionDecision() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamDecision.xlsx");
            var decisions = new List<(string Operation, int Items, ExecutionMode Mode)>();

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 1, "Value");
                    document.Save();
                }

                var options = new ExcelReadOptions();
                options.Execution.OperationThresholds["ReadRangeStream"] = 1;
                options.Execution.OnDecision = (operation, items, mode) => decisions.Add((operation, items, mode));

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var chunk = Assert.Single(reader.GetSheet("Data").ReadRangeStream("A1:A2"));

                Assert.Equal(2, chunk.RowCount);
                var decision = Assert.Single(decisions);
                Assert.Equal("ReadRangeStream", decision.Operation);
                Assert.Equal(2, decision.Items);
                Assert.Equal(ExecutionMode.Parallel, decision.Mode);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadObjectsAs_ReportsOwnExecutionDecision() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderObjectsAsDecision.xlsx");
            var decisions = new List<(string Operation, int Items, ExecutionMode Mode)>();

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Id");
                    sheet.CellValue(1, 2, "Name");
                    sheet.CellValue(2, 1, 1);
                    sheet.CellValue(2, 2, "One");
                    document.Save();
                }

                var options = new ExcelReadOptions();
                options.Execution.OperationThresholds["ReadObjectsAs"] = 1;
                options.Execution.OnDecision = (operation, items, mode) => decisions.Add((operation, items, mode));

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects<ReaderDecisionRecord>("A1:B2"));

                Assert.Equal(1, row.Id);
                Assert.Equal("One", row.Name);
                var decision = Assert.Single(decisions);
                Assert.Equal("ReadObjectsAs", decision.Operation);
                Assert.Equal(4, decision.Items);
                Assert.Equal(ExecutionMode.Parallel, decision.Mode);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_SequentialMode_ReturnsOrderedChunks() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamSequentialChunks.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "A");
                    sheet.CellValue(1, 2, 10);
                    sheet.CellValue(2, 1, "B");
                    sheet.CellValue(2, 2, 20);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var chunks = reader.GetSheet("Data")
                    .ReadRangeStream("A1:B2", chunkRows: 1, mode: ExecutionMode.Sequential)
                    .ToList();

                Assert.Equal(2, chunks.Count);
                Assert.Equal(1, chunks[0].StartRow);
                Assert.Equal(1, chunks[0].RowCount);
                Assert.Equal("A", chunks[0].Rows[0][0]);
                Assert.Equal(10D, chunks[0].Rows[0][1]);
                Assert.Equal(2, chunks[1].StartRow);
                Assert.Equal(1, chunks[1].RowCount);
                Assert.Equal("B", chunks[1].Rows[0][0]);
                Assert.Equal(20D, chunks[1].Rows[0][1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_ParallelMode_ReturnsOrderedChunks() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamParallelChunks.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    for (int row = 1; row <= 5; row++) {
                        sheet.CellValue(row, 1, $"R{row}");
                        sheet.CellValue(row, 2, row * 10D);
                    }
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var chunks = reader.GetSheet("Data")
                    .ReadRangeStream("A1:B5", chunkRows: 2, mode: ExecutionMode.Parallel)
                    .ToList();

                Assert.Equal(3, chunks.Count);
                Assert.Equal(new[] { 1, 3, 5 }, chunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal(new[] { 2, 2, 1 }, chunks.Select(chunk => chunk.RowCount).ToArray());
                Assert.Equal("R1", chunks[0].Rows[0][0]);
                Assert.Equal(20D, chunks[0].Rows[1][1]);
                Assert.Equal("R3", chunks[1].Rows[0][0]);
                Assert.Equal(40D, chunks[1].Rows[1][1]);
                Assert.Equal("R5", chunks[2].Rows[0][0]);
                Assert.Equal(50D, chunks[2].Rows[0][1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_HandlesOutOfOrderRowsWithinChunk() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamOutOfOrderRowsWithinChunk.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 1, "InRange");
                    sheet.CellValue(3, 1, "Tail");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 2U);

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                var sequentialChunk = Assert.Single(sheetReader
                    .ReadRangeStream("A1:A3", chunkRows: 3, mode: ExecutionMode.Sequential));
                Assert.Equal(1, sequentialChunk.StartRow);
                Assert.Equal(3, sequentialChunk.RowCount);
                Assert.Equal("Header", sequentialChunk.Rows[0][0]);
                Assert.Equal("InRange", sequentialChunk.Rows[1][0]);
                Assert.Equal("Tail", sequentialChunk.Rows[2][0]);

                var parallelChunk = Assert.Single(sheetReader
                    .ReadRangeStream("A1:A3", chunkRows: 3, mode: ExecutionMode.Parallel));
                Assert.Equal(1, parallelChunk.StartRow);
                Assert.Equal(3, parallelChunk.RowCount);
                Assert.Equal("Header", parallelChunk.Rows[0][0]);
                Assert.Equal("InRange", parallelChunk.Rows[1][0]);
                Assert.Equal("Tail", parallelChunk.Rows[2][0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_HandlesOutOfOrderRowsAcrossChunks() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamOutOfOrderRowsAcrossChunks.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "One");
                    sheet.CellValue(2, 1, "Two");
                    sheet.CellValue(3, 1, "Three");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 2U);

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                var sequentialChunks = sheetReader
                    .ReadRangeStream("A1:A3", chunkRows: 1, mode: ExecutionMode.Sequential)
                    .ToList();
                Assert.Equal(new[] { 1, 2, 3 }, sequentialChunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal(new[] { "One", "Two", "Three" }, sequentialChunks.Select(chunk => (string?)chunk.Rows[0][0]).ToArray());

                var parallelChunks = sheetReader
                    .ReadRangeStream("A1:A3", chunkRows: 1, mode: ExecutionMode.Parallel)
                    .ToList();
                Assert.Equal(new[] { 1, 2, 3 }, parallelChunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal(new[] { "One", "Two", "Three" }, parallelChunks.Select(chunk => (string?)chunk.Rows[0][0]).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_LargeOutOfOrderRangeUsesFallbackOrdering() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamLargeOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "One");
                    sheet.CellValue(2049, 1, "Middle");
                    sheet.CellValue(4097, 1, "Last");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 2049U);

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                var sequentialChunks = sheetReader
                    .ReadRangeStream("A1:A4097", chunkRows: 2048, mode: ExecutionMode.Sequential)
                    .ToList();
                Assert.Equal(new[] { 1, 2049, 4097 }, sequentialChunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal(new[] { "One", "Middle", "Last" }, sequentialChunks.Select(chunk => (string?)chunk.Rows[0][0]).ToArray());

                var parallelChunks = sheetReader
                    .ReadRangeStream("A1:A4097", chunkRows: 2048, mode: ExecutionMode.Parallel)
                    .ToList();
                Assert.Equal(new[] { 1, 2049, 4097 }, parallelChunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal(new[] { "One", "Middle", "Last" }, parallelChunks.Select(chunk => (string?)chunk.Rows[0][0]).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_AutomaticModeKeepsLargeOutOfOrderRowsOrdered() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamAutomaticLargeOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "One");
                    sheet.CellValue(2049, 1, "Middle");
                    sheet.CellValue(4097, 1, "Last");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 2049U);

                using var reader = ExcelDocumentReader.Open(filePath);
                var chunks = reader.GetSheet("Data")
                    .ReadRangeStream("A1:A4097", chunkRows: 2048)
                    .ToList();

                Assert.Equal(new[] { 1, 2049, 4097 }, chunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal(new[] { "One", "Middle", "Last" }, chunks.Select(chunk => (string?)chunk.Rows[0][0]).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeStream_KeepsSparseChunksBoundedByChunkRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRangeStreamSparseChunks.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First");
                    sheet.CellValue(10, 1, "Tenth");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var chunks = reader.GetSheet("Data")
                    .ReadRangeStream("A1:A10", chunkRows: 3, mode: ExecutionMode.Parallel)
                    .ToList();

                Assert.Equal(2, chunks.Count);
                Assert.Equal(new[] { 1, 10 }, chunks.Select(chunk => chunk.StartRow).ToArray());
                Assert.Equal(new[] { 3, 1 }, chunks.Select(chunk => chunk.RowCount).ToArray());
                Assert.Equal("First", chunks[0].Rows[0][0]);
                Assert.Null(chunks[0].Rows[1][0]);
                Assert.Null(chunks[0].Rows[2][0]);
                Assert.Equal("Tenth", chunks[1].Rows[0][0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private sealed class ReaderDecisionRecord {
            public int Id { get; set; }

            public string Name { get; set; } = string.Empty;
        }
    }
}
