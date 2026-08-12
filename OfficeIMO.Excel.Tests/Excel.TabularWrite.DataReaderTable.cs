using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void WriteDataReader_CompactTableStreamsAndRoundTrips() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Created", typeof(DateTime));
            table.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 8, 30, 0));
            table.Rows.Add(2, "Beta", new DateTime(2026, 8, 11, 9, 45, 0));
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions {
                    SheetName = "Query Results",
                    CreateTable = true,
                    TableName = "QueryResults",
                    TableStyle = ExcelTableStyle.TableStyleMedium9,
                    IncludeAutoFilter = false,
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            Assert.False(reader.IsClosed);
            Assert.Equal("Query Results", result.SheetName);
            Assert.Equal("QueryResults", result.TableName);
            Assert.Equal("A1:C3", result.Range);
            Assert.Equal(2, result.RowCount);

            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            Assert.Null(spreadsheet.WorkbookPart!.SharedStringTablePart);
            var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.Single();
            Assert.Single(worksheetPart.TableDefinitionParts);
            var savedTable = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("QueryResults", savedTable.Name?.Value);
            Assert.Equal("A1:C3", savedTable.Reference?.Value);
            Assert.Null(savedTable.AutoFilter);
            Assert.Equal("TableStyleMedium9", savedTable.TableStyleInfo?.Name?.Value);

            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToArray();
            Assert.All(cells.Take(3), cell => Assert.NotNull(cell.CellReference));
            Assert.All(cells.Skip(3), cell => Assert.Null(cell.CellReference));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));

            output.Position = 0;
            using var workbookReader = ExcelDocumentReader.Open(output);
            object?[,] values = workbookReader.GetSheet("Query Results").ReadRange("A1:C3");
            Assert.Equal("Alpha", values[1, 1]);
            Assert.Equal(new DateTime(2026, 8, 11, 9, 45, 0), values[2, 2]);
        }

        [Fact]
        public void WriteDataReader_CompactTableWithHeadersAndNoRowsIsValid() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions {
                    CreateTable = true,
                    TableName = "EmptyResults",
                    IncludeAutoFilter = true,
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            Assert.Equal("A1:B1", result.Range);
            Assert.Equal(0, result.RowCount);
            Assert.Equal("EmptyResults", result.TableName);

            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            var savedTable = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:B1", savedTable.Reference?.Value);
            Assert.Equal("A1:B1", savedTable.AutoFilter?.Reference?.Value);
            Assert.Equal(2U, savedTable.TableColumns?.Count?.Value);
            Assert.Equal(2, worksheetPart.Worksheet.Descendants<Cell>().Count());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteDataReader_CompactTableSupportsNonSeekableDestination() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add(1, "Alpha");
            using var reader = table.CreateDataReader();
            using var output = new NonSeekableWriteStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions {
                    CreateTable = true,
                    TableName = "StreamedResults",
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            Assert.Equal("A1:B2", result.Range);
            using var package = new MemoryStream(output.ToArray(), writable: false);
            using var spreadsheet = SpreadsheetDocument.Open(package, false);
            var savedTable = spreadsheet.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
            Assert.Equal("StreamedResults", savedTable.Name?.Value);
            Assert.Equal("A1:B2", savedTable.Reference?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteDataReader_CompactTableNormalizesPackageNames() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Rows.Add(1);
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();
            string requestedTableName = new string('T', 300);

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions {
                    SheetName = "'[]:*?/ '",
                    CreateTable = true,
                    TableName = requestedTableName,
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            Assert.Equal("Sheet1", result.SheetName);
            Assert.Equal(255, result.TableName?.Length);

            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            var savedSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single();
            var savedTable = spreadsheet.WorkbookPart.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
            Assert.Equal("Sheet1", savedSheet.Name?.Value);
            Assert.Equal(requestedTableName.Substring(0, 255), savedTable.Name?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteDataReader_CompactWriterTruncatesLongWorksheetName() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Rows.Add(1);
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions {
                    SheetName = new string('S', 40),
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            Assert.Equal(new string('S', 31), result.SheetName);
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            var savedSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single();
            Assert.Equal(new string('S', 31), savedSheet.Name?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteDataReader_DefaultTableNameMatchesBufferedFallback() {
            var table = new DataTable("SourceRows");
            table.Columns.Add("Id", typeof(int));
            table.Rows.Add(1);
            using var compactReader = table.CreateDataReader();
            using var bufferedReader = table.CreateDataReader();
            using var compactOutput = new MemoryStream();
            using var bufferedOutput = new MemoryStream();

            ExcelDataSetImportResult compactResult = ExcelDocument.WriteDataReader(
                compactOutput,
                compactReader,
                new ExcelTabularWriteOptions {
                    SheetName = "Data",
                    CreateTable = true,
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });
            ExcelDataSetImportResult bufferedResult = ExcelDocument.WriteDataReader(
                bufferedOutput,
                bufferedReader,
                new ExcelTabularWriteOptions {
                    SheetName = "Data",
                    CreateTable = true,
                    UseSharedStrings = true
                });

            Assert.Equal("ReaderData", compactResult.TableName);
            Assert.Equal(compactResult.TableName, bufferedResult.TableName);
            using var compactPackage = SpreadsheetDocument.Open(compactOutput, false);
            using var bufferedPackage = SpreadsheetDocument.Open(bufferedOutput, false);
            string? compactName = compactPackage.WorkbookPart!.WorksheetParts.Single()
                .TableDefinitionParts.Single().Table?.Name?.Value;
            string? bufferedName = bufferedPackage.WorkbookPart!.WorksheetParts.Single()
                .TableDefinitionParts.Single().Table?.Name?.Value;
            Assert.Equal("ReaderData", compactName);
            Assert.Equal(compactName, bufferedName);
            Assert.Empty(new OpenXmlValidator().Validate(compactPackage));
            Assert.Empty(new OpenXmlValidator().Validate(bufferedPackage));
        }
    }
}
