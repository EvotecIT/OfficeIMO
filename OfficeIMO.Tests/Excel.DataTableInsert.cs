using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for inserting DataTable content with mixed null values.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_InsertDataTable_BlanksMaintainNumericAndDateTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableNulls.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                var table = new DataTable();
                table.Columns.Add("Id", typeof(int));
                table.Columns.Add("Amount", typeof(double));
                table.Columns.Add("Date", typeof(DateTime));

                table.Rows.Add(1, 10.5, new DateTime(2024, 1, 1));

                var second = table.NewRow();
                second["Id"] = 2;
                second["Amount"] = DBNull.Value;
                second["Date"] = new DateTime(2024, 1, 2);
                table.Rows.Add(second);

                var third = table.NewRow();
                third["Id"] = 3;
                third["Amount"] = 5.75;
                third["Date"] = DBNull.Value;
                table.Rows.Add(third);

                sheet.InsertDataTable(table, includeHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = worksheetPart.Worksheet.Descendants<Cell>().ToList();

                Cell GetCell(string reference) {
                    return cells.First(c => c.CellReference == reference);
                }

                var amountRow2 = GetCell("B2");
                Assert.Equal(CellValues.Number, amountRow2.DataType!.Value);
                Assert.Equal(10.5.ToString(CultureInfo.InvariantCulture), amountRow2.CellValue!.Text);

                var dateRow2 = GetCell("C2");
                Assert.Equal(CellValues.Number, dateRow2.DataType!.Value);
                Assert.Equal(new DateTime(2024, 1, 1).ToOADate().ToString(CultureInfo.InvariantCulture), dateRow2.CellValue!.Text);

                var amountRow3 = GetCell("B3");
                Assert.Equal(CellValues.String, amountRow3.DataType!.Value);
                Assert.True(string.IsNullOrEmpty(amountRow3.CellValue!.Text));

                var dateRow3 = GetCell("C3");
                Assert.Equal(CellValues.Number, dateRow3.DataType!.Value);
                Assert.Equal(new DateTime(2024, 1, 2).ToOADate().ToString(CultureInfo.InvariantCulture), dateRow3.CellValue!.Text);

                var amountRow4 = GetCell("B4");
                Assert.Equal(CellValues.Number, amountRow4.DataType!.Value);
                Assert.Equal(5.75.ToString(CultureInfo.InvariantCulture), amountRow4.CellValue!.Text);

                var dateRow4 = GetCell("C4");
                Assert.Equal(CellValues.String, dateRow4.DataType!.Value);
                Assert.True(string.IsNullOrEmpty(dateRow4.CellValue!.Text));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_InsertDataTable_TimeSpanColumnGetsDurationFormat() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableDurations.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Durations");

                var table = new DataTable();
                table.Columns.Add("Task", typeof(string));
                table.Columns.Add("Elapsed", typeof(TimeSpan));

                table.Rows.Add("Build", TimeSpan.FromMinutes(90));
                table.Rows.Add("QA", new TimeSpan(2, 15, 30));

                sheet.InsertDataTable(table, includeHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var worksheetPart = workbookPart.WorksheetParts.First();

                Cell GetCell(string reference) => worksheetPart.Worksheet.Descendants<Cell>()
                    .First(c => c.CellReference == reference);

                var durationCell = GetCell("B2");
                Assert.True(durationCell.DataType == null || durationCell.DataType.Value == CellValues.Number);
                Assert.Equal(TimeSpan.FromMinutes(90).TotalDays.ToString(CultureInfo.InvariantCulture), durationCell.CellValue!.Text);

                var stylesPart = workbookPart.WorkbookStylesPart;
                Assert.NotNull(stylesPart);

                var numberingFormats = stylesPart!.Stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
                    .Where(n => n.FormatCode != null)
                    .ToList();
                Assert.NotNull(numberingFormats);

                var durationFormat = numberingFormats!.FirstOrDefault(n => string.Equals(n.FormatCode!.Value, "[h]:mm:ss", StringComparison.Ordinal));
                Assert.NotNull(durationFormat);

                uint numFmtId = durationFormat!.NumberFormatId!.Value;

                var cellFormats = stylesPart.Stylesheet!.CellFormats!.Elements<CellFormat>().ToList();
                int formatIndex = cellFormats.FindIndex(cf => cf.NumberFormatId != null && cf.NumberFormatId.Value == numFmtId && cf.ApplyNumberFormat?.Value == true);
                Assert.True(formatIndex >= 0, "Duration cell format should be registered.");

                Assert.NotNull(durationCell.StyleIndex);
                Assert.Equal((uint)formatIndex, durationCell.StyleIndex!.Value);

                var secondDuration = GetCell("B3");
                Assert.True(secondDuration.DataType == null || secondDuration.DataType.Value == CellValues.Number);
                Assert.Equal(new TimeSpan(2, 15, 30).TotalDays.ToString(CultureInfo.InvariantCulture), secondDuration.CellValue!.Text);
                Assert.Equal(durationCell.StyleIndex!.Value, secondDuration.StyleIndex!.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_ExtendsTableAndMapsColumnsByHeader() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTable.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                table.Rows.Add("EMEA", 120);

                Assert.Equal("A1:B3", sheet.InsertDataTableAsTable(table, tableName: "SalesTable"));

                var append = new DataTable();
                append.Columns.Add("Revenue", typeof(int));
                append.Columns.Add("Region", typeof(string));
                append.Rows.Add(150, "APAC");
                append.Rows.Add(175, "LATAM");

                Assert.Equal("A1:B5", sheet.AppendDataTableToTable(append, "SalesTable"));
                Assert.Equal("A1:B5", sheet.GetTableRange("SalesTable"));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = worksheetPart.TableDefinitionParts.First();
                Assert.Equal("A1:B5", tablePart.Table.Reference!.Value);
                Assert.Equal("A1:B5", tablePart.Table.GetFirstChild<AutoFilter>()!.Reference!.Value);

                Assert.Equal("APAC", GetCellText(spreadsheet, worksheetPart, "A4"));
                Assert.Equal("150", GetCellText(spreadsheet, worksheetPart, "B4"));
                Assert.Equal("LATAM", GetCellText(spreadsheet, worksheetPart, "A5"));
                Assert.Equal("175", GetCellText(spreadsheet, worksheetPart, "B5"));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
                Assert.Equal("A1:B5", document.Sheets[0].GetTableRange("SalesTable"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_ThrowsWhenColumnIsMissing() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableMissingColumn.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");

                var append = new DataTable();
                append.Columns.Add("Region", typeof(string));
                append.Columns.Add("Amount", typeof(int));
                append.Rows.Add("APAC", 150);

                var exception = Assert.Throws<ArgumentException>(() => sheet.AppendDataTableToTable(append, "SalesTable"));
                Assert.Contains("Revenue", exception.Message);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_ThrowsWhenCellsBelowTableContainData() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableOccupiedCells.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");

                sheet.CellValue(3, 1, "Existing");

                var append = new DataTable();
                append.Columns.Add("Region", typeof(string));
                append.Columns.Add("Revenue", typeof(int));
                append.Rows.Add("APAC", 150);

                var exception = Assert.Throws<InvalidOperationException>(() => sheet.AppendDataTableToTable(append, "SalesTable"));
                Assert.Contains("A3", exception.Message);
            }

            File.Delete(filePath);
        }
    }
}
