using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("Range")]
        [InlineData("DataTable")]
        [InlineData("DataSet")]
        [InlineData("StreamingReader")]
        public void HeaderOnlyTablesHaveNoRecordsAndAppendImmediatelyAfterHeader(string sourceKind) {
            string path = Path.Combine(_directoryWithFiles, "EmptyRecords." + sourceKind + ".xlsx");
            if (sourceKind == "StreamingReader") {
                using var source = EmptyTableRecordSource();
                using var reader = source.CreateDataReader();
                using var writer = ExcelDocument.CreateWriter(path, new ExcelWriterOptions {
                    CreateTable = true, TableName = "Records", UseSharedStrings = false
                });
                writer.WriteTable("Data", reader);
            } else {
                using var document = ExcelDocument.Create(path);
                if (sourceKind == "DataSet") {
                    using var dataSet = new DataSet();
                    dataSet.Tables.Add(EmptyTableRecordSource());
                    document.InsertDataSet(dataSet);
                } else {
                    var sheet = document.AddSheet("Data");
                    if (sourceKind == "DataTable") {
                        using var source = EmptyTableRecordSource();
                        sheet.AddTable(source, new ExcelTableOptions { Name = "Records" });
                    } else {
                        sheet.Cell(1, 1).SetValue("Id");
                        sheet.Cell(1, 2).SetValue("Name");
                        sheet.AddTable("A1:B1", name: "Records");
                    }
                    Assert.Empty(sheet.Table("Records").RowsAs<EmptyTableRecord>());
                }
                document.Save();
            }
            using (var package = SpreadsheetDocument.Open(path, false)) {
                var table = package.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
                Assert.Equal("A1:B1", table.Reference!.Value);
                Assert.Equal("A1:B1", table.AutoFilter!.Reference!.Value);
                Assert.NotEqual(true, table.InsertRow?.Value);
            }
            using (var loaded = ExcelDocument.Load(path)) {
                var sheet = loaded.Sheets[0];
                Assert.Empty(sheet.Table("Records").RowsAs<EmptyTableRecord>());
                Assert.Empty(sheet.TableToDataTable("Records").Rows.Cast<DataRow>());
                Assert.Empty(loaded.ValidateOpenXml());
            }
            using (var document = ExcelDocument.Load(path)) {
                var sheet = document.Sheets[0];
                using var rows = EmptyTableRecordSource();
                rows.Rows.Add(1, "Alpha");
                rows.Rows.Add(2, "Beta");
                Assert.Equal("A1:B3", sheet.AppendDataTableToTable(rows, "Records"));
                Assert.Equal(new[] { 1, 2 }, sheet.Table("Records").RowsAs<EmptyTableRecord>().Select(row => row.Id.GetValueOrDefault()));
                document.Save();
            }
            using (var package = SpreadsheetDocument.Open(path, false)) {
                var part = package.WorkbookPart!.WorksheetParts.Single();
                Assert.Equal("A1:B3", part.TableDefinitionParts.Single().Table!.Reference!.Value);
                Assert.NotEqual(true, part.TableDefinitionParts.Single().Table!.InsertRow?.Value);
                Assert.Equal("1", part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A2").CellValue!.Text);
                Assert.Equal("2", part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A3").CellValue!.Text);
            }
            using var saved = ExcelDocument.Load(path);
            Assert.Equal(2, saved.Sheets[0].Table("Records").RowsAs<EmptyTableRecord>().Count());
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void ExplicitBlankTableDataRowIsReadAndRetainedWhenAppending(bool fromDataTable, bool nativeInsertRow) {
            string path = Path.Combine(_directoryWithFiles, "ExplicitBlankRecord.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                if (fromDataTable) {
                    using var source = EmptyTableRecordSource();
                    source.Rows.Add(DBNull.Value, DBNull.Value);
                    sheet.AddTable(source, new ExcelTableOptions { Name = "Records" });
                } else {
                    sheet.Cell(1, 1).SetValue("Id");
                    sheet.Cell(1, 2).SetValue("Name");
                    sheet.AddTable("A1:B2", name: "Records");
                }
                document.Save();
            }
            if (nativeInsertRow) {
                using var imported = SpreadsheetDocument.Open(path, true);
                var table = imported.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
                table.InsertRow = true;
                table.Save();
            }
            using (var loaded = ExcelDocument.Load(path)) {
                Assert.Single(loaded.Sheets[0].TableToDataTable("Records").Rows.Cast<DataRow>());
                Assert.Single(loaded.Sheets[0].Table("Records").RowsAs<EmptyTableRecord>());
            }
            using (var document = ExcelDocument.Load(path)) {
                using var rows = EmptyTableRecordSource();
                rows.Rows.Add(1, "Alpha");
                Assert.Equal("A1:B3", document.Sheets[0].AppendDataTableToTable(rows, "Records"));
                document.Save();
            }
            using var saved = ExcelDocument.Load(path);
            Assert.Equal(2, saved.Sheets[0].TableToDataTable("Records").Rows.Count);
            using var package = SpreadsheetDocument.Open(path, false);
            var part = package.WorkbookPart!.WorksheetParts.Single();
            Assert.Equal("A1:B3", part.TableDefinitionParts.Single().Table!.Reference!.Value);
            Assert.NotEqual(true, part.TableDefinitionParts.Single().Table!.InsertRow?.Value);
            Assert.Equal("1", part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A3").CellValue!.Text);
        }

        private static DataTable EmptyTableRecordSource() {
            var source = new DataTable("Records");
            source.Columns.Add("Id", typeof(int));
            source.Columns.Add("Name", typeof(string));
            return source;
        }

        public sealed class EmptyTableRecord {
            public int? Id { get; set; }
            public string? Name { get; set; }
        }
    }
}
