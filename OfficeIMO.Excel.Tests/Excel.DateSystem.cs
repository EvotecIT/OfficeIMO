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
    public partial class Excel {
        private const double Excel1904DateOffsetDays = 1462d;

        [Fact]
        public void DateSystem_1904_WritesWorkbookFlagAndAdjustedSerials() {
            string filePath = Path.Combine(_directoryWithFiles, "DateSystem1904.xlsx");
            var date = new DateTime(2024, 1, 1);
            var minimum = new DateTime(2024, 1, 1);
            var maximum = new DateTime(2024, 12, 31);

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    document.DateSystem = ExcelDateSystem.NineteenFour;
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, date);
                    sheet.ValidationDate("A1:A10", DataValidationOperatorValues.Between, minimum, maximum);
                    document.Save();
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                    AssertWorkbookUses1904DateSystem(spreadsheet);

                    WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    string serialText = GetCellValueText(worksheetPart, "A1");
                    Assert.Equal(Expected1904Serial(date), double.Parse(serialText, CultureInfo.InvariantCulture), 6);

                    DataValidation validation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
                    Assert.Equal(Expected1904Serial(minimum), double.Parse(validation.GetFirstChild<Formula1>()!.Text, CultureInfo.InvariantCulture), 6);
                    Assert.Equal(Expected1904Serial(maximum), double.Parse(validation.GetFirstChild<Formula2>()!.Text, CultureInfo.InvariantCulture), 6);
                }

                using (var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { TreatDatesUsingNumberFormat = true })) {
                    var cell = reader.GetSheet("Data").EnumerateCells().Single(item => item.Row == 1 && item.Column == 1);
                    DateTime readDate = Assert.IsType<DateTime>(cell.Value);
                    Assert.Equal(date, readDate);
                }

                using (ExcelDocument loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    Assert.Equal(ExcelDateSystem.NineteenFour, loaded.DateSystem);
                    Assert.Equal(ExcelDateSystem.NineteenFour, loaded.CreateInspectionSnapshot().DateSystem);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DirectDataSet_1904_WritesWorkbookFlagAndAdjustedDateSerials() {
            var date = new DateTime(2024, 2, 3);
            using var stream = new MemoryStream();
            var table = new DataTable("Data");
            table.Columns.Add("When", typeof(DateTime));
            table.Rows.Add(date);
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(stream, dataSet, dateSystem: ExcelDateSystem.NineteenFour);
            stream.Position = 0;

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false);
            AssertWorkbookUses1904DateSystem(spreadsheet);

            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            string serialText = GetCellValueText(worksheetPart, "A2");
            Assert.Equal(Expected1904Serial(date), double.Parse(serialText, CultureInfo.InvariantCulture), 6);
        }

        [Fact]
        public void DateSystem_ChangingExistingWorkbook_ConvertsDateStyledSerials() {
            string filePath = Path.Combine(_directoryWithFiles, "DateSystemChangeConvertsSerials.xlsx");
            var date = new DateTime(2024, 3, 4);

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    ExcelSheet sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, date);
                    document.Save();
                }

                using (var document = ExcelDocument.Load(filePath)) {
                    document.DateSystem = ExcelDateSystem.NineteenFour;
                    document.Save();
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                    AssertWorkbookUses1904DateSystem(spreadsheet);
                    WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    string serialText = GetCellValueText(worksheetPart, "A1");
                    Assert.Equal(Expected1904Serial(date), double.Parse(serialText, CultureInfo.InvariantCulture), 6);
                }

                using (var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { TreatDatesUsingNumberFormat = true })) {
                    var cell = reader.GetSheet("Data").EnumerateCells().Single(item => item.Row == 1 && item.Column == 1);
                    Assert.Equal(date, Assert.IsType<DateTime>(cell.Value));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static double Expected1904Serial(DateTime value)
            => value.ToOADate() - Excel1904DateOffsetDays;

        private static string GetCellValueText(WorksheetPart worksheetPart, string cellReference) {
            Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference == cellReference);
            return cell.CellValue?.Text ?? string.Empty;
        }

        private static void AssertWorkbookUses1904DateSystem(SpreadsheetDocument spreadsheet) {
            WorkbookProperties? properties = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<WorkbookProperties>();
            Assert.NotNull(properties);
            Assert.True(properties!.Date1904?.Value);
        }
    }
}
