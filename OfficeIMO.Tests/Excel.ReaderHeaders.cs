using System;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private sealed class DuplicateHeaderRow {
            public string? Value { get; set; }
            public string? Value_2 { get; set; }
        }

        private sealed class FriendlyHeaderRow {
            public string? FirstName { get; set; }
            public string? FirstName_2 { get; set; }
            public int TotalAmount2 { get; set; }
        }

        private sealed class ExactFriendlyPrecedenceRow {
            public string? FirstName { get; set; }
        }

        private sealed class AttributedHeaderRow {
            [DisplayName("First Name")]
            public string? GivenName { get; set; }

            [DataMember(Name = "Status Code")]
            public string? Status { get; set; }

            [ExcelColumn("Total %", "Total Percent")]
            public int CompletionPercent { get; set; }
        }

        private sealed class AmbiguousAttributedHeaderRow {
            [ExcelColumn("Status")]
            public string? PrimaryStatus { get; set; }

            [DisplayName("Status")]
            public string? SecondaryStatus { get; set; }
        }

        private sealed class StrictMappedRow {
            public string? Name { get; set; }
        }

        private sealed class NullableTypedRow {
            public int? Score { get; set; }
            public bool? Active { get; set; }
            public DateTime? CreatedOn { get; set; }
            public double? Amount { get; set; }
        }

        private sealed class DecimalTypedRow {
            public decimal Amount { get; set; }
            public decimal? OptionalAmount { get; set; }
        }

        private sealed class CultureDoubleTypedRow {
            public double Amount { get; set; }
        }

        private sealed class DateStyledNumericTypedRow {
            public double NumericValue { get; set; }
            public DateTime DateValue { get; set; }
            public string? TextValue { get; set; }
        }

        private static void AssertRangeEqual(object?[,] expected, object?[,] actual) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(expected.GetLength(1), actual.GetLength(1));

            for (int r = 0; r < expected.GetLength(0); r++) {
                for (int c = 0; c < expected.GetLength(1); c++) {
                    Assert.Equal(expected[r, c], actual[r, c]);
                }
            }
        }

        private static void AssertDataTablesEqual(DataTable expected, DataTable actual) {
            Assert.Equal(expected.Columns.Count, actual.Columns.Count);
            Assert.Equal(expected.Rows.Count, actual.Rows.Count);

            for (int c = 0; c < expected.Columns.Count; c++) {
                Assert.Equal(expected.Columns[c].ColumnName, actual.Columns[c].ColumnName);
            }

            for (int r = 0; r < expected.Rows.Count; r++) {
                for (int c = 0; c < expected.Columns.Count; c++) {
                    Assert.Equal(expected.Rows[r][c], actual.Rows[r][c]);
                }
            }
        }

        [Fact]
        public void Reader_ReadObjects_DisambiguatesDuplicateAndNormalizedHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDuplicateHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(1, 2, "  Value  ");
                    sheet.CellValue(1, 3, "");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, "Gamma");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));

                Assert.Equal("Alpha", row["Value"]);
                Assert.Equal("Beta", row["Value_2"]);
                Assert.Equal("Gamma", row["Column3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_DisambiguatesDuplicateHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDuplicateHeadersDataTable.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(1, 2, "Status");
                    sheet.CellValue(1, 3, " status ");
                    sheet.CellValue(2, 1, "OK");
                    sheet.CellValue(2, 2, "Warning");
                    sheet.CellValue(2, 3, "Error");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Status", "Status_2", "status_3" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("OK", table.Rows[0]["Status"]);
                Assert.Equal("Warning", table.Rows[0]["Status_2"]);
                Assert.Equal("Error", table.Rows[0]["status_3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRange_ForcedExecutionModesReturnSameValues() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderExecutionModes.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.CellValue(1, 3, "Active");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 12.5d);
                    sheet.CellValue(2, 3, true);
                    sheet.CellValue(3, 1, "Beta");
                    sheet.CellValue(3, 2, 25d);
                    sheet.CellValue(3, 3, false);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                object?[,] automatic = sheetReader.ReadRange("A1:C3");
                object?[,] sequential = sheetReader.ReadRange("A1:C3", ExecutionMode.Sequential);
                object?[,] parallel = sheetReader.ReadRange("A1:C3", ExecutionMode.Parallel);

                AssertRangeEqual(automatic, sequential);
                AssertRangeEqual(automatic, parallel);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_ForcedExecutionModesReturnSameValues() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableExecutionModes.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.CellValue(1, 3, "Active");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 12.5d);
                    sheet.CellValue(3, 1, "Beta");
                    sheet.CellValue(3, 3, true);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                DataTable automatic = sheetReader.ReadRangeAsDataTable("A1:C3");
                DataTable sequential = sheetReader.ReadRangeAsDataTable("A1:C3", mode: ExecutionMode.Sequential);
                DataTable parallel = sheetReader.ReadRangeAsDataTable("A1:C3", mode: ExecutionMode.Parallel);

                AssertDataTablesEqual(automatic, sequential);
                AssertDataTablesEqual(automatic, parallel);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_WithoutHeadersPreservesAllRowsAndBlankCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableNoHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Alpha");
                    sheet.CellValue(1, 3, 10);
                    sheet.CellValue(2, 2, "Beta");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2", headersInFirstRow: false);

                Assert.Equal(new[] { "Column1", "Column2", "Column3" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal("Alpha", table.Rows[0]["Column1"]);
                Assert.Equal(DBNull.Value, table.Rows[0]["Column2"]);
                Assert.Equal(10d, table.Rows[0]["Column3"]);
                Assert.Equal(DBNull.Value, table.Rows[1]["Column1"]);
                Assert.Equal("Beta", table.Rows[1]["Column2"]);
                Assert.Equal(DBNull.Value, table.Rows[1]["Column3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_InfersStableColumnTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableTypedColumns.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.CellValue(1, 3, "Created");
                    sheet.CellValue(1, 4, "Active");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 12.5d);
                    sheet.CellValue(2, 3, new DateTime(2024, 1, 2));
                    sheet.CellValue(2, 4, true);
                    sheet.CellValue(3, 1, "Beta");
                    sheet.CellValue(3, 2, 25d);
                    sheet.CellValue(3, 3, new DateTime(2024, 1, 3));
                    sheet.CellValue(3, 4, false);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:D3");

                Assert.Equal(typeof(string), table.Columns["Name"]!.DataType);
                Assert.Equal(typeof(double), table.Columns["Amount"]!.DataType);
                Assert.Equal(typeof(DateTime), table.Columns["Created"]!.DataType);
                Assert.Equal(typeof(bool), table.Columns["Active"]!.DataType);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_MapsMemoryPackageWithInferredTypes() {
            var expectedDate = new DateTime(2024, 1, 2);
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Created");
                sheet.CellValue(1, 4, "Active");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(2, 2, 12.5d);
                sheet.CellValue(2, 3, expectedDate);
                sheet.CellValue(2, 4, true);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:D2");

            Assert.Equal(typeof(string), table.Columns["Name"]!.DataType);
            Assert.Equal(typeof(double), table.Columns["Amount"]!.DataType);
            Assert.Equal(typeof(DateTime), table.Columns["Created"]!.DataType);
            Assert.Equal(typeof(bool), table.Columns["Active"]!.DataType);
            DataRow row = Assert.Single(table.Rows.Cast<DataRow>());
            Assert.Equal("Alpha", row["Name"]);
            Assert.Equal(12.5d, row["Amount"]);
            Assert.Equal(expectedDate, row["Created"]);
            Assert.Equal(true, row["Active"]);
        }

        [Fact]
        public void Reader_ReadColumn_MapsMemoryPackageWithWideRows() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(1, 3, "Amount");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Alpha");
                sheet.CellValue(2, 3, 12.5d);
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 2, "Beta");
                sheet.CellValue(3, 3, 25d);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            var values = reader.GetSheet("Data").ReadColumn("A1:A3").ToList();

            Assert.Equal("Id", values[0]);
            Assert.Equal(1, Convert.ToInt32(values[1], CultureInfo.InvariantCulture));
            Assert.Equal(2, Convert.ToInt32(values[2], CultureInfo.InvariantCulture));
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_KeepsMixedColumnObjectTyped() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableMixedColumn.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Mixed");
                    sheet.CellValue(2, 1, 10d);
                    sheet.CellValue(3, 1, "Ten");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:A3");

                Assert.Equal(typeof(object), table.Columns["Mixed"]!.DataType);
                Assert.Equal(10d, table.Rows[0]["Mixed"]);
                Assert.Equal("Ten", table.Rows[1]["Mixed"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_ReportsOwnExecutionDecision() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataTableDecision.xlsx");
            var decisions = new List<(string Operation, int Items, ExecutionMode Mode)>();

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Value");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 1);
                    document.Save();
                }

                var options = new ExcelReadOptions();
                options.Execution.OperationThresholds["ReadRangeAsDataTable"] = 1;
                options.Execution.OnDecision = (operation, items, mode) => decisions.Add((operation, items, mode));

                using var reader = ExcelDocumentReader.Open(filePath, options);
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:B2");

                Assert.Single(table.Rows);
                var decision = Assert.Single(decisions);
                Assert.Equal("ReadRangeAsDataTable", decision.Operation);
                Assert.Equal(4, decision.Items);
                Assert.Equal(ExecutionMode.Parallel, decision.Mode);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_ReportsOwnExecutionDecision() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsDecision.xlsx");
            var decisions = new List<(string Operation, int Items, ExecutionMode Mode)>();

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Ignored");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Value");
                    document.Save();
                }

                var options = new ExcelReadOptions();
                options.Execution.OperationThresholds["ReadObjectsAs"] = 1;
                options.Execution.OnDecision = (operation, items, mode) => decisions.Add((operation, items, mode));

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects<StrictMappedRow>("A1:B2"));

                Assert.Equal("Alpha", row.Name);
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
        public void Reader_ReadRange_PreservesRichSharedAndInlineStrings() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRichStrings.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Shared");
                    sheet.CellValue(1, 2, "Inline");
                    sheet.CellValue(2, 1, "placeholder");
                    sheet.CellValue(2, 2, "placeholder");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var row = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 2);
                    var sharedCell = row.Elements<Cell>().First(c => c.CellReference?.Value == "A2");
                    var inlineCell = row.Elements<Cell>().First(c => c.CellReference?.Value == "B2");

                    int sharedIndex = int.Parse(sharedCell.CellValue!.Text);
                    var sharedTable = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable!;
                    sharedTable.ReplaceChild(
                        new SharedStringItem(
                            new Run(new Text("Shared ")),
                            new Run(new Text("Rich"))),
                        sharedTable.Elements<SharedStringItem>().ElementAt(sharedIndex));

                    inlineCell.CellValue = null;
                    inlineCell.DataType = CellValues.InlineString;
                    inlineCell.InlineString = new InlineString(
                        new Run(new Text("Inline ")),
                        new Run(new Text("Rich")));

                    sharedTable.Save();
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Data").ReadRange("A2:B2");

                Assert.Equal("Shared Rich", values[0, 0]);
                Assert.Equal("Inline Rich", values[0, 1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRange_IgnoresSharedStringPhoneticRuns() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderPhoneticSharedStrings.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "placeholder");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference?.Value == "A1");
                    int sharedIndex = int.Parse(cell.CellValue!.Text);
                    var sharedTable = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable!;
                    var item = sharedTable.Elements<SharedStringItem>().ElementAt(sharedIndex);
                    item.InnerXml = "<r><t>Displayed</t></r><rPh sb=\"0\" eb=\"9\"><t>Phonetic</t></rPh>";
                    sharedTable.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:A1");

                Assert.Equal("Displayed", values[0, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeSequential_DoesNotStopAtOutOfOrderRowsBeyondRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderOutOfOrderRows.xlsx");

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
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var row2 = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 2U);
                    row2.Remove();
                    sheetData.Append(row2);
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:A2", ExecutionMode.Sequential);

                Assert.Equal("Header", values[0, 0]);
                Assert.Equal("InRange", values[1, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_DoesNotStopAtOutOfOrderRowsBeyondRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "InRange");
                    sheet.CellValue(5, 1, "Outside");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var row2 = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 2U);
                    row2.Remove();
                    sheetData.Append(row2);
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects<StrictMappedRow>("A1:A2", ExecutionMode.Sequential));

                Assert.Equal("InRange", row.Name);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_DoesNotStopAtOutOfOrderRowsBeyondRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedStreamOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "InRange");
                    sheet.CellValue(5, 1, "Outside");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var row2 = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 2U);
                    row2.Remove();
                    sheetData.Append(row2);
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<StrictMappedRow>("A1:A2"));

                Assert.Equal("InRange", row.Name);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_UsesLateOutOfOrderHeaderRow() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedStreamLateHeaderRow.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "InRange");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var headerRow = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 1U);
                    headerRow.Remove();
                    sheetData.Append(headerRow);
                    worksheetPart.Worksheet.Save();
                }

                var options = new ExcelReadOptions { StrictTypedMapping = true };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<StrictMappedRow>("A1:A2"));

                Assert.Equal("InRange", row.Name);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_PreservesOutOfOrderRowsInsideRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedStreamOutOfOrderInsideRange.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "Second");
                    sheet.CellValue(3, 1, "Third");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var row2 = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 2U);
                    row2.Remove();
                    sheetData.Append(row2);
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadObjectsStream<StrictMappedRow>("A1:A3").ToList();

                Assert.Equal(2, rows.Count);
                Assert.Equal("Second", rows[0].Name);
                Assert.Equal("Third", rows[1].Name);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRange_IgnoresMalformedCellReferenceWithoutRowNumber() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderMalformedCellReference.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "ShouldIgnore");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference?.Value == "A1");
                    cell.CellReference = "A";
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:A1", ExecutionMode.Sequential);

                Assert.Null(values[0, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_Rows_CanBeEnumeratedAfterReaderScopeDisposes() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowsMaterialized.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, 12.5d);
                    document.Save();
                }

                IEnumerable<Dictionary<string, object?>> rows;
                using (var document = ExcelDocument.Load(filePath)) {
                    rows = document.GetSheet("Data").Rows("A1:B2");
                }

                var row = Assert.Single(rows);
                Assert.Equal("Alpha", row["Name"]);
                Assert.Equal(12.5d, row["Amount"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_RowsAs_CanBeEnumeratedAfterReaderScopeDisposes() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowsAsMaterialized.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "First Name");
                    sheet.CellValue(1, 3, "Total Amount 2");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, 42);
                    document.Save();
                }

                IEnumerable<FriendlyHeaderRow> rows;
                using (var document = ExcelDocument.Load(filePath)) {
                    rows = document.GetSheet("Data").RowsAs<FriendlyHeaderRow>("A1:C2");
                }

                var row = Assert.Single(rows);
                Assert.Equal("Alpha", row.FirstName);
                Assert.Equal("Beta", row.FirstName_2);
                Assert.Equal(42, row.TotalAmount2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_RowsAsStream_EnumeratesWhileDocumentScopeIsOpen() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderRowsAsStreamBridge.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "First Name");
                    sheet.CellValue(1, 3, "Total Amount 2");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, 42);
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var row = Assert.Single(loadedDocument.GetSheet("Data").RowsAsStream<FriendlyHeaderRow>("A1:C2"));

                Assert.Equal("Alpha", row.FirstName);
                Assert.Equal("Beta", row.FirstName_2);
                Assert.Equal(42, row.TotalAmount2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Fluent_AsObjectsStream_EnumeratesWhileDocumentScopeIsOpen() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFluentObjectsStreamBridge.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "First Name");
                    sheet.CellValue(1, 3, "Total Amount 2");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, 42);
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var row = Assert.Single(loadedDocument.Read().Sheet("Data").Range("A1:C2").AsObjectsStream<FriendlyHeaderRow>());

                Assert.Equal("Alpha", row.FirstName);
                Assert.Equal("Beta", row.FirstName_2);
                Assert.Equal(42, row.TotalAmount2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_ReadHelpers_ExposeDisambiguatedHeadersConsistently() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDuplicateHeadersEditable.xlsx");

            try {
                using (var createdDocument = ExcelDocument.Create(filePath)) {
                    var createdSheet = createdDocument.AddWorkSheet("Data");
                    createdSheet.CellValue(1, 1, "Value");
                    createdSheet.CellValue(1, 2, "Value");
                    createdSheet.CellValue(2, 1, "Left");
                    createdSheet.CellValue(2, 2, "Right");
                    createdDocument.Save();
                }

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");

                var map = sheet.GetHeaderMap();
                Assert.Equal(1, map["Value"]);
                Assert.Equal(2, map["Value_2"]);

                var editable = Assert.Single(sheet.RowsObjects("A1:B2"));
                Assert.Equal("Left", editable["Value"].Value);
                Assert.Equal("Right", editable["Value_2"].Value);

                var typed = Assert.Single(sheet.RowsAs<DuplicateHeaderRow>("A1:B2"));
                Assert.Equal("Left", typed.Value);
                Assert.Equal("Right", typed.Value_2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_BlankGeneratedHeaders_DoNotStealExplicitColumnNames() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderBlankGeneratedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "");
                    sheet.CellValue(1, 2, "Column1");
                    sheet.CellValue(2, 1, "Generated");
                    sheet.CellValue(2, 2, "Explicit");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:B2"));

                Assert.Equal("Generated", row["Column1_2"]);
                Assert.Equal("Explicit", row["Column1"]);

                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:B2");
                Assert.Equal(new[] { "Column1_2", "Column1" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ExplicitSuffixedHeaders_RemainStableWhenBaseHeaderRepeats() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderExplicitSuffixedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(1, 2, "Value_2");
                    sheet.CellValue(1, 3, "Value");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, "Gamma");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));

                Assert.Equal("Alpha", row["Value"]);
                Assert.Equal("Beta", row["Value_2"]);
                Assert.Equal("Gamma", row["Value_3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_BlankGeneratedHeaders_DoNotStealExplicitSuffixedNames() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderBlankGeneratedHeadersReservedSuffix.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "");
                    sheet.CellValue(1, 2, "Column1");
                    sheet.CellValue(1, 3, "Column1_2");
                    sheet.CellValue(2, 1, "Generated");
                    sheet.CellValue(2, 2, "ExplicitBase");
                    sheet.CellValue(2, 3, "ExplicitSuffix");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Column1_3", "Column1", "Column1_2" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("Generated", row["Column1_3"]);
                Assert.Equal("ExplicitBase", row["Column1"]);
                Assert.Equal("ExplicitSuffix", row["Column1_2"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMapCache_RebuildsAfterHeaderRenameWithinSameUsedRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheRename.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(1, 2, "Value");
                    sheet.CellValue(2, 1, "Open");
                    sheet.CellValue(2, 2, "10");
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var initialMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, initialMap["Status"]);

                loadedSheet.CellValue(1, 1, "State");

                var refreshedMap = loadedSheet.GetHeaderMap();
                Assert.False(refreshedMap.ContainsKey("Status"));
                Assert.Equal(1, refreshedMap["State"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMap_ReturnedMapDoesNotMutateCachedLookup() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheDefensiveCopy.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(1, 2, "Value");
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var map = loadedSheet.GetHeaderMap();
                map["Status"] = 99;
                map["Injected"] = 3;

                Assert.True(loadedSheet.TryGetColumnIndexByHeader("Status", out int statusColumn));
                Assert.Equal(1, statusColumn);
                Assert.False(loadedSheet.TryGetColumnIndexByHeader("Injected", out _));

                var secondMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, secondMap["Status"]);
                Assert.False(secondMap.ContainsKey("Injected"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMapCache_RebuildsWhenUsedRangeShiftsAfterWrite() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheUsedRangeShift.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "Region");
                    sheet.CellValue(2, 3, "Amount");
                    sheet.CellValue(3, 2, "North");
                    sheet.CellValue(3, 3, 10);
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var initialMap = loadedSheet.GetHeaderMap();
                Assert.Equal(2, initialMap["Region"]);
                Assert.Equal(3, initialMap["Amount"]);

                loadedSheet.CellValue(1, 1, "Id");
                loadedSheet.CellValue(1, 2, "Region");
                loadedSheet.CellValue(1, 3, "Amount");

                var refreshedMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, refreshedMap["Id"]);
                Assert.Equal(2, refreshedMap["Region"]);
                Assert.Equal(3, refreshedMap["Amount"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMap_ReadsOnlyFirstUsedRow() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderMapFirstUsedRowOnly.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Id");
                    sheet.CellValue(1, 2, "Value");
                    for (int row = 2; row <= 1000; row++) {
                        sheet.CellValue(row, 1, (double)(row - 1));
                        sheet.CellValue(row, 2, 10d);
                    }

                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");
                var options = new ExcelReadOptions {
                    CellValueConverter = context => {
                        if (context.TypeHint is null && context.RawText == "10") {
                            throw new InvalidOperationException("Header map should not read data rows.");
                        }

                        return ExcelCellValue.NotHandled;
                    }
                };

                var map = loadedSheet.GetHeaderMap(options);

                Assert.Equal(1, map["Id"]);
                Assert.Equal(2, map["Value"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMapCache_RebuildsAfterParallelCellValuesOverwriteHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheParallelCellValues.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "OldName");
                    sheet.CellValue(1, 2, "OldValue");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "1");
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var initialMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, initialMap["OldName"]);
                Assert.Equal(2, initialMap["OldValue"]);

                loadedSheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Value")
                }, ExecutionMode.Parallel);

                var refreshedMap = loadedSheet.GetHeaderMap();
                Assert.False(refreshedMap.ContainsKey("OldName"));
                Assert.False(refreshedMap.ContainsKey("OldValue"));
                Assert.Equal(1, refreshedMap["Name"]);
                Assert.Equal(2, refreshedMap["Value"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMapCache_RebuildsAfterParallelInsertDataTableOverwriteHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheParallelDataTable.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "OldStatus");
                    sheet.CellValue(1, 2, "OldOwner");
                    sheet.CellValue(2, 1, "Open");
                    sheet.CellValue(2, 2, "Alice");
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var initialMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, initialMap["OldStatus"]);
                Assert.Equal(2, initialMap["OldOwner"]);

                var table = new DataTable();
                table.Columns.Add("Status");
                table.Columns.Add("Owner");
                table.Rows.Add("Closed", "Bob");

                loadedSheet.InsertDataTable(table, startRow: 1, startColumn: 1, includeHeaders: true, mode: ExecutionMode.Parallel);

                var refreshedMap = loadedSheet.GetHeaderMap();
                Assert.False(refreshedMap.ContainsKey("OldStatus"));
                Assert.False(refreshedMap.ContainsKey("OldOwner"));
                Assert.Equal(1, refreshedMap["Status"]);
                Assert.Equal(2, refreshedMap["Owner"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ShiftedRange_DisambiguatesBlankAndExplicitHeadersLocally() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderShiftedRangeHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "");
                    sheet.CellValue(2, 3, "Column1");
                    sheet.CellValue(2, 4, "Column1_2");
                    sheet.CellValue(3, 2, "Generated");
                    sheet.CellValue(3, 3, "ExplicitBase");
                    sheet.CellValue(3, 4, "ExplicitSuffix");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("B2:D3"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("B2:D3");

                Assert.Equal(new[] { "Column1_3", "Column1", "Column1_2" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("Generated", row["Column1_3"]);
                Assert.Equal("ExplicitBase", row["Column1"]);
                Assert.Equal("ExplicitSuffix", row["Column1_2"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_NormalizeHeadersFalse_PreservesWhitespaceDistinctHeadersAcrossReadSurfaces() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderNormalizeHeadersFalse.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(1, 2, "  Value  ");
                    sheet.CellValue(1, 3, "");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, "Gamma");
                    document.Save();
                }

                var options = new ExcelReadOptions { NormalizeHeaders = false };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Value", "  Value  ", "Column3" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("Alpha", row["Value"]);
                Assert.Equal("Beta", row["  Value  "]);
                Assert.Equal("Gamma", row["Column3"]);

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");
                var headerMap = loadedSheet.GetHeaderMap(options);
                var editable = Assert.Single(loadedSheet.RowsObjects("A1:C2", options));

                Assert.Equal(1, headerMap["Value"]);
                Assert.Equal(2, headerMap["  Value  "]);
                Assert.Equal(3, headerMap["Column3"]);
                Assert.Equal("Alpha", editable["Value"].Value);
                Assert.Equal("Beta", editable["  Value  "].Value);
                Assert.Equal("Gamma", editable["Column3"].Value);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_MapFriendlyDuplicateHeadersToDisambiguatedProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFriendlyTypedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "First Name");
                    sheet.CellValue(1, 3, "Total Amount 2");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, "Bob");
                    sheet.CellValue(2, 3, 42);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var typedFromReader = Assert.Single(reader.GetSheet("Data").ReadObjects<FriendlyHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromReader.FirstName);
                Assert.Equal("Bob", typedFromReader.FirstName_2);
                Assert.Equal(42, typedFromReader.TotalAmount2);

                using var loadedDocument = ExcelDocument.Load(filePath);
                var typedFromSheet = Assert.Single(loadedDocument.GetSheet("Data").RowsAs<FriendlyHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromSheet.FirstName);
                Assert.Equal("Bob", typedFromSheet.FirstName_2);
                Assert.Equal(42, typedFromSheet.TotalAmount2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_PreferExactPropertyHeadersOverEarlierFriendlyMatches() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFriendlyTypedHeadersExactPrecedence.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "FirstName");
                    sheet.CellValue(2, 1, "AliasValue");
                    sheet.CellValue(2, 2, "ExactValue");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var typed = Assert.Single(reader.GetSheet("Data").ReadObjects<ExactFriendlyPrecedenceRow>("A1:B2"));

                Assert.Equal("ExactValue", typed.FirstName);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_ParsesDoubleTextWithConfiguredCultureFirst() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderCultureDoubleTyped.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Amount");
                    sheet.CellValue(2, 1, "1,23");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions {
                    Culture = CultureInfo.GetCultureInfo("pl-PL")
                });
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects<CultureDoubleTypedRow>("A1:A2"));

                Assert.Equal(1.23d, row.Amount, precision: 10);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_MapAttributeBasedHeaderAliases() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderAttributedTypedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "Status Code");
                    sheet.CellValue(1, 3, "Total %");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, "OK");
                    sheet.CellValue(2, 3, 97);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var typedFromReader = Assert.Single(reader.GetSheet("Data").ReadObjects<AttributedHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromReader.GivenName);
                Assert.Equal("OK", typedFromReader.Status);
                Assert.Equal(97, typedFromReader.CompletionPercent);

                using var loadedDocument = ExcelDocument.Load(filePath);
                var typedFromSheet = Assert.Single(loadedDocument.GetSheet("Data").RowsAs<AttributedHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromSheet.GivenName);
                Assert.Equal("OK", typedFromSheet.Status);
                Assert.Equal(97, typedFromSheet.CompletionPercent);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_MapNullableValueTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderNullableTypedHeaders.xlsx");
            var expectedDate = new DateTime(2024, 5, 12, 9, 30, 0);

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(1, 2, "Active");
                    sheet.CellValue(1, 3, "CreatedOn");
                    sheet.CellValue(1, 4, "Amount");
                    sheet.CellValue(2, 1, 42);
                    sheet.CellValue(2, 2, true);
                    sheet.CellValue(2, 3, expectedDate);
                    sheet.CellValue(2, 4, 123.45);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadObjects<NullableTypedRow>("A1:D3").ToList();

                Assert.Equal(2, rows.Count);
                Assert.Equal(42, rows[0].Score);
                Assert.True(rows[0].Active);
                Assert.Equal(expectedDate, rows[0].CreatedOn);
                Assert.Equal(123.45, rows[0].Amount);
                Assert.Null(rows[1].Score);
                Assert.Null(rows[1].Active);
                Assert.Null(rows[1].CreatedOn);
                Assert.Null(rows[1].Amount);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_HandleOutOfOrderCellsWithinWideRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedOutOfOrderCellsWithinWideRows.xlsx");
            var expectedDate = new DateTime(2024, 5, 12, 9, 30, 0);

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(1, 2, "Active");
                    sheet.CellValue(1, 3, "CreatedOn");
                    sheet.CellValue(1, 4, "Amount");
                    sheet.CellValue(1, 5, "Ignored");
                    sheet.CellValue(2, 1, 42);
                    sheet.CellValue(2, 2, true);
                    sheet.CellValue(2, 3, expectedDate);
                    sheet.CellValue(2, 4, 123.45);
                    sheet.CellValue(2, 5, "tail");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var row = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Single(r => r.RowIndex?.Value == 2U);
                    var cells = row.Elements<Cell>().ToDictionary(c => c.CellReference!.Value!);
                    row.RemoveAllChildren<Cell>();
                    row.Append(cells["C2"]);
                    row.Append(cells["A2"]);
                    row.Append(cells["E2"]);
                    row.Append(cells["B2"]);
                    row.Append(cells["D2"]);
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                var rowFromMaterializedReader = Assert.Single(sheetReader.ReadObjects<NullableTypedRow>("A1:E2"));
                Assert.Equal(42, rowFromMaterializedReader.Score);
                Assert.True(rowFromMaterializedReader.Active);
                Assert.Equal(expectedDate, rowFromMaterializedReader.CreatedOn);
                Assert.Equal(123.45, rowFromMaterializedReader.Amount);

                var rowFromStreamingReader = Assert.Single(sheetReader.ReadObjectsStream<NullableTypedRow>("A1:E2"));
                Assert.Equal(42, rowFromStreamingReader.Score);
                Assert.True(rowFromStreamingReader.Active);
                Assert.Equal(expectedDate, rowFromStreamingReader.CreatedOn);
                Assert.Equal(123.45, rowFromStreamingReader.Amount);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_MapDecimalValueTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDecimalTypedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Amount");
                    sheet.CellValue(1, 2, "OptionalAmount");
                    sheet.CellValue(2, 1, 123.45m);
                    sheet.CellValue(2, 2, 678.90m);
                    sheet.CellValue(3, 1, 11.25m);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadObjects<DecimalTypedRow>("A1:B3").ToList();

                Assert.Equal(2, rows.Count);
                Assert.Equal(123.45m, rows[0].Amount);
                Assert.Equal(678.90m, rows[0].OptionalAmount);
                Assert.Equal(11.25m, rows[1].Amount);
                Assert.Null(rows[1].OptionalAmount);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_ParallelKeepsDateStyledNumericTargetsNumeric() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDateStyledNumericTypedHeaders.xlsx");
            const double serialValue = 1.5d;

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "NumericValue");
                    sheet.CellValue(1, 2, "DateValue");
                    sheet.CellValue(1, 3, "TextValue");
                    sheet.CellValue(2, 1, serialValue);
                    sheet.CellValue(2, 2, serialValue);
                    sheet.CellValue(2, 3, serialValue);
                    sheet.ColumnStyleByHeader("NumericValue").NumberFormat("[h]:mm");
                    sheet.ColumnStyleByHeader("DateValue").NumberFormat("[h]:mm");
                    sheet.ColumnStyleByHeader("TextValue").NumberFormat("[h]:mm");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects<DateStyledNumericTypedRow>("A1:C2", ExecutionMode.Parallel));

                Assert.Equal(serialValue, row.NumericValue);
                Assert.Equal(DateTime.FromOADate(serialValue), row.DateValue);
                Assert.Equal(DateTime.FromOADate(serialValue).ToString(CultureInfo.InvariantCulture), row.TextValue);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_ParallelHonorsHandledNullTypeConverterForDateStyledNumeric() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDateStyledNumericTypeConverter.xlsx");
            const double serialValue = 1.5d;

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "NumericValue");
                    sheet.CellValue(2, 1, serialValue);
                    sheet.ColumnStyleByHeader("NumericValue").NumberFormat("[h]:mm");
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    TypeConverter = (value, targetType, culture) =>
                        targetType == typeof(double) ? (true, null) : (false, null)
                };

                using var sequentialReader = ExcelDocumentReader.Open(filePath, options);
                using var parallelReader = ExcelDocumentReader.Open(filePath, options);
                var sequentialRow = Assert.Single(sequentialReader.GetSheet("Data").ReadObjects<DateStyledNumericTypedRow>("A1:A2", ExecutionMode.Sequential));
                var row = Assert.Single(parallelReader.GetSheet("Data").ReadObjects<DateStyledNumericTypedRow>("A1:A2", ExecutionMode.Parallel));

                Assert.Equal(sequentialRow.NumericValue, row.NumericValue);
                Assert.Equal(0d, row.NumericValue);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_ParallelHonorsHandledCellConverterForDateStyledNumeric() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDateStyledNumericCellConverter.xlsx");
            const double serialValue = 1.5d;

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "NumericValue");
                    sheet.CellValue(2, 1, serialValue);
                    sheet.ColumnStyleByHeader("NumericValue").NumberFormat("[h]:mm");
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    CellValueConverter = context =>
                        context.RawText == serialValue.ToString(CultureInfo.InvariantCulture)
                            ? new ExcelCellValue("not-a-number")
                            : ExcelCellValue.NotHandled
                };

                using var sequentialReader = ExcelDocumentReader.Open(filePath, options);
                using var parallelReader = ExcelDocumentReader.Open(filePath, options);
                var sequentialRow = Assert.Single(sequentialReader.GetSheet("Data").ReadObjects<DateStyledNumericTypedRow>("A1:A2", ExecutionMode.Sequential));
                var row = Assert.Single(parallelReader.GetSheet("Data").ReadObjects<DateStyledNumericTypedRow>("A1:A2", ExecutionMode.Parallel));

                Assert.Equal(sequentialRow.NumericValue, row.NumericValue);
                Assert.Equal(0d, row.NumericValue);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_MapsRowsWithoutMaterializingAllObjects() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamHeaders.xlsx");
            var expectedDate = new DateTime(2024, 3, 2);

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(1, 2, "Active");
                    sheet.CellValue(1, 3, "CreatedOn");
                    sheet.CellValue(1, 4, "Amount");
                    sheet.CellValue(2, 1, 42);
                    sheet.CellValue(2, 2, true);
                    sheet.CellValue(2, 3, expectedDate);
                    sheet.CellValue(2, 4, 123.45);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<NullableTypedRow>("A1:D2"));

                Assert.Equal(42, row.Score);
                Assert.True(row.Active);
                Assert.Equal(expectedDate, row.CreatedOn);
                Assert.Equal(123.45, row.Amount);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_HonorsCellValueConverter() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamCellConverter.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(2, 1, 42);
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    CellValueConverter = context => context.RawText == "42" ? new ExcelCellValue("100") : ExcelCellValue.NotHandled
                };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<NullableTypedRow>("A1:A2"));

                Assert.Equal(100, row.Score);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_ParsesDoubleTextWithConfiguredCultureFirst() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamCultureDoubleTyped.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Amount");
                    sheet.CellValue(2, 1, "1,23");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions {
                    Culture = CultureInfo.GetCultureInfo("pl-PL")
                });
                var row = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<CultureDoubleTypedRow>("A1:A2"));

                Assert.Equal(1.23d, row.Amount, precision: 10);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_HonorsCellValueConverter() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsCellConverter.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(2, 1, 42);
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    CellValueConverter = context => context.RawText == "42" ? new ExcelCellValue("100") : ExcelCellValue.NotHandled
                };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects<NullableTypedRow>("A1:A2", ExecutionMode.Sequential));

                Assert.Equal(100, row.Score);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_MapsRowsFromMemoryPackage() {
            var expectedDate = new DateTime(2024, 3, 2);
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Score");
                sheet.CellValue(1, 2, "Active");
                sheet.CellValue(1, 3, "CreatedOn");
                sheet.CellValue(1, 4, "Amount");
                sheet.CellValue(2, 1, 42);
                sheet.CellValue(2, 2, true);
                sheet.CellValue(2, 3, expectedDate);
                sheet.CellValue(2, 4, 123.45);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            var row = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<NullableTypedRow>("A1:D2"));

            Assert.Equal(42, row.Score);
            Assert.True(row.Active);
            Assert.Equal(expectedDate, row.CreatedOn);
            Assert.Equal(123.45, row.Amount);
        }

        [Fact]
        public void Reader_TypedObjects_AutomaticUsesSinglePassForMemoryPackage() {
            var expectedDate = new DateTime(2024, 3, 2);
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Score");
                sheet.CellValue(1, 2, "Active");
                sheet.CellValue(1, 3, "CreatedOn");
                sheet.CellValue(1, 4, "Amount");
                sheet.CellValue(2, 1, 42);
                sheet.CellValue(2, 2, true);
                sheet.CellValue(2, 3, expectedDate);
                sheet.CellValue(2, 4, 123.45);
            }

            var options = new ExcelReadOptions();
            options.Execution.OperationThresholds["ReadObjectsAs"] = 1;

            using var reader = ExcelDocumentReader.Open(memory.ToArray(), options);
            var row = Assert.Single(reader.GetSheet("Data").ReadObjects<NullableTypedRow>("A1:D2"));

            Assert.Equal(42, row.Score);
            Assert.True(row.Active);
            Assert.Equal(expectedDate, row.CreatedOn);
            Assert.Equal(123.45, row.Amount);
        }

        [Fact]
        public void Reader_TypedObjectsStream_ReturnsEmptyWhenRangeContainsOnlyHeader() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamHeaderOnly.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadObjectsStream<NullableTypedRow>("A1:A1").ToList();

                Assert.Empty(rows);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_PreservesBlankRowsInsideRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamBlankRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(2, 1, 42);
                    sheet.CellValue(4, 1, 44);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadObjectsStream<NullableTypedRow>("A1:A4").ToList();

                Assert.Equal(3, rows.Count);
                Assert.Equal(42, rows[0].Score);
                Assert.Null(rows[1].Score);
                Assert.Equal(44, rows[2].Score);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_HonorsHandledConverters() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamConverters.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(2, 1, 42);
                    document.Save();
                }

                var options = new ExcelReadOptions {
                    TypeConverter = (value, targetType, culture) =>
                        targetType == typeof(int) ? (true, 7) : (false, null)
                };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<NullableTypedRow>("A1:A2"));

                Assert.Equal(7, row.Score);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_KeepsDateStyledNumericTargetsNumeric() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamDateStyledNumeric.xlsx");
            const double serialValue = 1.5d;

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "NumericValue");
                    sheet.CellValue(1, 2, "DateValue");
                    sheet.CellValue(1, 3, "TextValue");
                    sheet.CellValue(2, 1, serialValue);
                    sheet.CellValue(2, 2, serialValue);
                    sheet.CellValue(2, 3, serialValue);
                    sheet.ColumnStyleByHeader("NumericValue").NumberFormat("[h]:mm");
                    sheet.ColumnStyleByHeader("DateValue").NumberFormat("[h]:mm");
                    sheet.ColumnStyleByHeader("TextValue").NumberFormat("[h]:mm");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<DateStyledNumericTypedRow>("A1:C2"));

                Assert.Equal(serialValue, row.NumericValue);
                Assert.Equal(DateTime.FromOADate(serialValue), row.DateValue);
                Assert.Equal(DateTime.FromOADate(serialValue).ToString(CultureInfo.InvariantCulture), row.TextValue);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_ThrowsWhenCancellationArrivesDuringMaterialization() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsMaterializationCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(2, 1, 42);
                    sheet.CellValue(3, 1, 43);
                    document.Save();
                }

                using var cts = new CancellationTokenSource();
                var options = new ExcelReadOptions {
                    TypeConverter = (value, targetType, culture) => {
                        cts.Cancel();
                        return (false, null);
                    }
                };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadObjects<NullableTypedRow>("A1:A3", ct: cts.Token).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_ThrowsWhenCancellationArrivesDuringEnumeration() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(2, 1, 42);
                    sheet.CellValue(3, 1, 43);
                    document.Save();
                }

                using var cts = new CancellationTokenSource();
                var options = new ExcelReadOptions {
                    TypeConverter = (value, targetType, culture) => {
                        cts.Cancel();
                        return (false, null);
                    }
                };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadObjectsStream<NullableTypedRow>("A1:A3", cts.Token).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjectsStream_ThrowsWhenCancellationArrivesOnFinalConversion() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedObjectsStreamFinalConversionCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Score");
                    sheet.CellValue(2, 1, 42);
                    document.Save();
                }

                using var cts = new CancellationTokenSource();
                var options = new ExcelReadOptions {
                    TypeConverter = (value, targetType, culture) => {
                        cts.Cancel();
                        return (false, null);
                    }
                };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadObjectsStream<NullableTypedRow>("A1:A2", cts.Token).ToList());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAs_ThrowsWhenCancellationArrivesDuringConversion() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderTypedRangeConversionCancellation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 42);
                    sheet.CellValue(1, 2, 43);
                    document.Save();
                }

                using var cts = new CancellationTokenSource();
                var options = new ExcelReadOptions {
                    TypeConverter = (value, targetType, culture) => {
                        cts.Cancel();
                        return (false, null);
                    }
                };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                Assert.Throws<OperationCanceledException>(() =>
                    reader.GetSheet("Data").ReadRangeAs<int>("A1:B1", ct: cts.Token));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_ReportAmbiguousAliasDiagnostics() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderAmbiguousTypedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(2, 1, "Open");
                    document.Save();
                }

                var diagnostics = new List<string>();
                var options = new ExcelReadOptions();
                options.Execution.OnInfo = diagnostics.Add;

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var typed = Assert.Single(reader.GetSheet("Data").ReadObjects<AmbiguousAttributedHeaderRow>("A1:A2"));

                Assert.Null(typed.PrimaryStatus);
                Assert.Null(typed.SecondaryStatus);
                Assert.Contains(diagnostics, message => message.Contains("TypedRead AmbiguousMapping", StringComparison.Ordinal)
                    && message.Contains("AmbiguousAttributedHeaderRow", StringComparison.Ordinal)
                    && message.Contains("Status", StringComparison.Ordinal));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_StrictMapping_ThrowsOnUnmappedHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderStrictTypedHeaders.Unmapped.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Status");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, "Open");
                    document.Save();
                }

                var options = new ExcelReadOptions { StrictTypedMapping = true };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var exception = Assert.Throws<InvalidOperationException>(() => reader.GetSheet("Data").ReadObjects<StrictMappedRow>("A1:B2").ToList());

                Assert.Contains("strict", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("UnmappedHeader", exception.Message, StringComparison.Ordinal);
                Assert.Contains("Status", exception.Message, StringComparison.Ordinal);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_StrictMapping_ThrowsOnAmbiguousAliases() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderStrictTypedHeaders.Ambiguous.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(2, 1, "Open");
                    document.Save();
                }

                var options = new ExcelReadOptions { StrictTypedMapping = true };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var exception = Assert.Throws<InvalidOperationException>(() => reader.GetSheet("Data").ReadObjects<AmbiguousAttributedHeaderRow>("A1:A2").ToList());

                Assert.Contains("strict", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("AmbiguousMapping", exception.Message, StringComparison.Ordinal);
                Assert.Contains("Status", exception.Message, StringComparison.Ordinal);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
