using System;
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
        public void Reader_RowReaders_HandleLargeSortedSparseRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowReadersLargeSortedSparseRanges.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
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
    }
}
