using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for various data validation types.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void ValidationWholeNumberBetween() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationWholeNumber.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.ValidationWholeNumber("A1:A10", DataValidationOperatorValues.Between, 1, 10, errorTitle: "Error", errorMessage: "1-10");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.Whole, dv.Type!.Value);
                Assert.Equal(DataValidationOperatorValues.Between, dv.Operator!.Value);
                Assert.Equal("1", dv.GetFirstChild<Formula1>()!.Text);
                Assert.Equal("10", dv.GetFirstChild<Formula2>()!.Text);
                Assert.Equal("Error", dv.ErrorTitle!.Value);
                Assert.Equal("1-10", dv.Error!.Value);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void ValidationDecimalGreaterThan() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationDecimal.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.ValidationDecimal("B1:B10", DataValidationOperatorValues.GreaterThan, 5.5);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.Decimal, dv.Type!.Value);
                Assert.Equal(DataValidationOperatorValues.GreaterThan, dv.Operator!.Value);
                Assert.Equal("5.5", dv.GetFirstChild<Formula1>()!.Text);
                Assert.Null(dv.GetFirstChild<Formula2>());
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void ValidationDateLessThanOrEqual() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationDate.xlsx");
            DateTime dt = new DateTime(2024, 1, 1);
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.ValidationDate("C1:C10", DataValidationOperatorValues.LessThanOrEqual, dt);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.Date, dv.Type!.Value);
                Assert.Equal(DataValidationOperatorValues.LessThanOrEqual, dv.Operator!.Value);
                string expected = dt.ToOADate().ToString(CultureInfo.InvariantCulture);
                Assert.Equal(expected, dv.GetFirstChild<Formula1>()!.Text);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void ValidationTimeEqual() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationTime.xlsx");
            TimeSpan ts = TimeSpan.FromHours(12);
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.ValidationTime("D1:D10", DataValidationOperatorValues.Equal, ts);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.Time, dv.Type!.Value);
                Assert.Equal(DataValidationOperatorValues.Equal, dv.Operator!.Value);
                string expected = ts.TotalDays.ToString(CultureInfo.InvariantCulture);
                Assert.Equal(expected, dv.GetFirstChild<Formula1>()!.Text);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void ValidationTextLengthLessThan() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationTextLength.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.ValidationTextLength("E1:E10", DataValidationOperatorValues.LessThan, 10);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.TextLength, dv.Type!.Value);
                Assert.Equal(DataValidationOperatorValues.LessThan, dv.Operator!.Value);
                Assert.Equal("10", dv.GetFirstChild<Formula1>()!.Text);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void ValidationCustomFormula() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationCustom.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.ValidationCustomFormula("F1:F10", "SUM(A1:B1)>10");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.Custom, dv.Type!.Value);
                Assert.Null(dv.Operator);
                Assert.Equal("SUM(A1:B1)>10", dv.GetFirstChild<Formula1>()!.Text);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void ValidationListNamedRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationListNamedRange.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet options = document.AddWorkSheet("Options");
                ExcelSheet data = document.AddWorkSheet("Data");

                options.CellValue(1, 1, "Open");
                options.CellValue(2, 1, "Closed");
                options.CellValue(3, 1, "Pending");
                document.SetNamedRange("StatusOptions", "'Options'!A1:A3", save: false);
                data.ValidationListNamedRange("B2:B5", "StatusOptions");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.Last();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.List, dv.Type!.Value);
                Assert.Equal("=StatusOptions", dv.GetFirstChild<Formula1>()!.Text);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void ValidationListRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationListRange.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet options = document.AddWorkSheet("Options");
                ExcelSheet data = document.AddWorkSheet("Data");

                options.CellValue(1, 1, "Open");
                options.CellValue(2, 1, "Closed");
                options.CellValue(3, 1, "Pending");
                data.ValidationListRange("B2:B5", "A1:A3", "Options");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.Last();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.List, dv.Type!.Value);
                Assert.Equal("='Options'!A1:A3", dv.GetFirstChild<Formula1>()!.Text);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void ValidationListRangeOnCurrentSheet() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidationListRangeOnCurrentSheet.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet data = document.AddWorkSheet("Data");

                data.CellValue(1, 4, "Open");
                data.CellValue(2, 4, "Closed");
                data.CellValue(3, 4, "Pending");
                data.ValidationListRange("B2:B5", "D1:D3");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                DataValidation dv = wsPart.Worksheet.Descendants<DataValidation>().First();
                Assert.Equal(DataValidationValues.List, dv.Type!.Value);
                Assert.Equal("=D1:D3", dv.GetFirstChild<Formula1>()!.Text);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}

