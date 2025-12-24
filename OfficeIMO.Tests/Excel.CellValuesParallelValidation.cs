using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void WorksheetValidationHonorsDiagnosticsConfiguration() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetValidation.Diagnostics.xlsx");

            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorkSheet("Data");

            var cells = Enumerable.Range(1, 10).Select(i => (i, 1, (object)$"Value {i}"));
            sheet.CellValues(cells, ExecutionMode.Parallel);

            var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var firstCell = sheetData.Elements<Row>().First().Elements<Cell>().First();
            firstCell.CellReference = "A99";

            var disabled = new ExecutionPolicy { WorksheetValidation = WorksheetValidationMode.Disabled };
            WorksheetIntegrityValidator.Validate(sheet.WorksheetPart, disabled, sheet.Name);

            var diagnosticsOnly = new ExecutionPolicy { WorksheetValidation = WorksheetValidationMode.DiagnosticsOnly };
            WorksheetIntegrityValidator.Validate(sheet.WorksheetPart, diagnosticsOnly, sheet.Name);

            diagnosticsOnly.DiagnosticsRequested = true;
            var ex = Assert.Throws<InvalidOperationException>(() => WorksheetIntegrityValidator.Validate(sheet.WorksheetPart, diagnosticsOnly, sheet.Name));
            Assert.Contains("does not match", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WorksheetValidationDetectsStructuralCorruption() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetValidation.Corruption.xlsx");

            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorkSheet("Data");

            var cells = Enumerable.Range(1, 5).SelectMany(row =>
                Enumerable.Range(1, 3).Select(col => (row, col, (object)$"R{row}C{col}")));
            sheet.CellValues(cells, ExecutionMode.Parallel);

            var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var secondRow = sheetData.Elements<Row>().ElementAt(1);
            secondRow.RowIndex = 1;

            var policy = new ExecutionPolicy { WorksheetValidation = WorksheetValidationMode.Always };
            var ex = Assert.Throws<InvalidOperationException>(() => WorksheetIntegrityValidator.Validate(sheet.WorksheetPart, policy, sheet.Name));
            Assert.Contains("non-increasing row indices", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WorksheetValidationIsFasterThanLegacyOuterXmlParsing() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetValidation.Performance.xlsx");

            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorkSheet("Data");

            var cells = Enumerable.Range(1, 1000)
                .SelectMany(row => Enumerable.Range(1, 8).Select(col => (row, col, (object)$"R{row}C{col}")))
                .ToList();

            sheet.CellValues(cells, ExecutionMode.Parallel);

            // Warm up targeted validation once before timing.
            WorksheetIntegrityValidator.Validate(sheet.WorksheetPart, new ExecutionPolicy {
                WorksheetValidation = WorksheetValidationMode.Always,
                DiagnosticsRequested = true,
            }, sheet.Name);

            var targeted = WorksheetIntegrityValidator.MeasureTargetedValidation(sheet.WorksheetPart, iterations: 3, sheet.Name);
            var legacy = WorksheetIntegrityValidator.MeasureLegacyOuterXml(sheet.WorksheetPart, iterations: 3);

            Assert.True(targeted < legacy, $"Targeted validation ({targeted}) should be faster than legacy OuterXml parsing ({legacy}).");
        }
    }
}
