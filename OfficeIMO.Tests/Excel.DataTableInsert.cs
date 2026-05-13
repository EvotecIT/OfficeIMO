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
        public void Test_AppendDataTableToTable_ThrowsWhenColumnCountDoesNotMatch() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableColumnCountMismatch.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");

                var append = new DataTable();
                append.Columns.Add("Region", typeof(string));
                append.Rows.Add("APAC");

                var exception = Assert.Throws<ArgumentException>(() => sheet.AppendDataTableToTable(append, "SalesTable"));
                Assert.Contains("1 columns", exception.Message);
                Assert.Contains("2 columns", exception.Message);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_HeaderlessTableUsesPositionalMapping() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableHeaderless.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);

                Assert.Equal("A1:B1", sheet.InsertDataTableAsTable(table, includeHeaders: false, tableName: "HeaderlessSales"));

                var append = new DataTable();
                append.Columns.Add("Region", typeof(string));
                append.Columns.Add("Revenue", typeof(int));
                append.Rows.Add("APAC", 150);
                append.Rows.Add("EMEA", 200);

                Assert.Equal("A1:B3", sheet.AppendDataTableToTable(append, "HeaderlessSales"));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = worksheetPart.TableDefinitionParts.First();
                Assert.Equal((uint)0, tablePart.Table.HeaderRowCount!.Value);
                Assert.Equal("A1:B3", tablePart.Table.Reference!.Value);

                Assert.Equal("APAC", GetCellText(spreadsheet, worksheetPart, "A2"));
                Assert.Equal("150", GetCellText(spreadsheet, worksheetPart, "B2"));
                Assert.Equal("EMEA", GetCellText(spreadsheet, worksheetPart, "A3"));
                Assert.Equal("200", GetCellText(spreadsheet, worksheetPart, "B3"));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_HiddenHeadersUseMatchingColumnNames() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableHiddenHeaders.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                sheet.InsertDataTableAsTable(table, tableName: "HiddenHeaderSales");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Table table = spreadsheet.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table;
                table.HeaderRowCount = 0U;
                table.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var append = new DataTable();
                append.Columns.Add("Revenue", typeof(int));
                append.Columns.Add("Region", typeof(string));
                append.Rows.Add(150, "APAC");

                Assert.Equal("A1:B3", document.Sheets[0].AppendDataTableToTable(append, "HiddenHeaderSales"));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = worksheetPart.TableDefinitionParts.First();
                Assert.Equal((uint)0, tablePart.Table.HeaderRowCount!.Value);
                Assert.Equal("A1:B3", tablePart.Table.Reference!.Value);
                Assert.Equal("APAC", GetCellText(spreadsheet, worksheetPart, "A3"));
                Assert.Equal("150", GetCellText(spreadsheet, worksheetPart, "B3"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_ColumnLikeHeadersUseHeaderMapping() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableColumnLikeHeaders.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, "Column1");
                sheet.CellValue(1, 2, "Column2");
                sheet.CellValue(2, 1, "first");
                sheet.CellValue(2, 2, "second");
                sheet.AddTable("A1:B2", true, "ColumnNamedSales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);

                var append = new DataTable();
                append.Columns.Add("Column2", typeof(string));
                append.Columns.Add("Column1", typeof(string));
                append.Rows.Add("fourth", "third");

                Assert.Equal("A1:B3", sheet.AppendDataTableToTable(append, "ColumnNamedSales"));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = worksheetPart.TableDefinitionParts.First();
                Assert.Equal("A1:B3", tablePart.Table.Reference!.Value);
                Assert.Equal("third", GetCellText(spreadsheet, worksheetPart, "A3"));
                Assert.Equal("fourth", GetCellText(spreadsheet, worksheetPart, "B3"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_EmptyTableKeepsExistingRange() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableEmpty.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);

                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");

                var append = new DataTable();
                append.Columns.Add("Revenue", typeof(int));
                append.Columns.Add("Region", typeof(string));

                Assert.Equal("A1:B2", sheet.AppendDataTableToTable(append, "SalesTable"));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = worksheetPart.TableDefinitionParts.First();
                Assert.Equal("A1:B2", tablePart.Table.Reference!.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_ThrowsWhenFormulaExistsBelowTable() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableFormulaBelow.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");

                sheet.CellFormula(3, 2, "SUM(B2:B2)");

                var append = new DataTable();
                append.Columns.Add("Region", typeof(string));
                append.Columns.Add("Revenue", typeof(int));
                append.Rows.Add("APAC", 150);

                var exception = Assert.Throws<InvalidOperationException>(() => sheet.AppendDataTableToTable(append, "SalesTable"));
                Assert.Contains("B3", exception.Message);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_AllowsHistoricalTotalsRowShownFlag() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableHistoricalTotals.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Table table = spreadsheet.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table;
                table.TotalsRowShown = true;
                table.TotalsRowCount = 0U;
                table.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var append = new DataTable();
                append.Columns.Add("Region", typeof(string));
                append.Columns.Add("Revenue", typeof(int));
                append.Rows.Add("APAC", 150);

                Assert.Equal("A1:B3", document.Sheets[0].AppendDataTableToTable(append, "SalesTable"));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = worksheetPart.TableDefinitionParts.First();
                Assert.Equal("A1:B3", tablePart.Table.Reference!.Value);
                Assert.Equal("APAC", GetCellText(spreadsheet, worksheetPart, "A3"));
                Assert.Equal("150", GetCellText(spreadsheet, worksheetPart, "B3"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_ThrowsWhenTotalsRowIsActive() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableActiveTotals.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Table table = spreadsheet.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table;
                table.TotalsRowShown = true;
                table.TotalsRowCount = 1U;
                table.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var append = new DataTable();
                append.Columns.Add("Region", typeof(string));
                append.Columns.Add("Revenue", typeof(int));
                append.Rows.Add("APAC", 150);

                var exception = Assert.Throws<InvalidOperationException>(() => document.Sheets[0].AppendDataTableToTable(append, "SalesTable"));
                Assert.Contains("totals row", exception.Message);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AppendDataTableToTable_ThrowsWhenTotalsRowShownWithoutCount() {
            string filePath = Path.Combine(_directoryWithFiles, "DataTableAppendTableTotalsShownNoCount.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");

                var table = new DataTable();
                table.Columns.Add("Region", typeof(string));
                table.Columns.Add("Revenue", typeof(int));
                table.Rows.Add("NA", 100);
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Table table = spreadsheet.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table;
                table.TotalsRowShown = true;
                table.TotalsRowCount = null;
                table.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var append = new DataTable();
                append.Columns.Add("Region", typeof(string));
                append.Columns.Add("Revenue", typeof(int));
                append.Rows.Add("APAC", 150);

                var exception = Assert.Throws<InvalidOperationException>(() => document.Sheets[0].AppendDataTableToTable(append, "SalesTable"));
                Assert.Contains("totals row", exception.Message);
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
