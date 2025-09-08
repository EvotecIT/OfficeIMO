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
        }
    }
}

