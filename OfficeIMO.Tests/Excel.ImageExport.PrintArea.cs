using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests {
    public class ExcelImageExportPrintAreaTests {
        [Fact]
        public void ExcelWorksheet_ImageExportUsesPrintAreaWhenRequested() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            sheet.CellValue(1, 1, "Outside");
            sheet.CellValue(2, 2, "North");
            sheet.CellValue(2, 3, 10);
            sheet.CellValue(3, 2, "South");
            sheet.CellValue(3, 3, 20);
            document.SetPrintArea(sheet, "B2:C3", save: false);

            ExcelRangeVisualSnapshot snapshot = sheet.CreateVisualSnapshot(new ExcelWorksheetImageExportOptions {
                UsePrintArea = true,
                ShowGridlines = false
            });
            OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                UsePrintArea = true,
                ShowGridlines = false
            });

            Assert.Equal("B2:C3", snapshot.Range);
            Assert.Equal(new[] { 2, 3 }, snapshot.Rows.Select(row => row.Index).ToArray());
            Assert.Equal(new[] { 2, 3 }, snapshot.Columns.Select(column => column.Index).ToArray());
            Assert.DoesNotContain(snapshot.Cells, cell => cell.Text == "Outside");
            Assert.Contains(snapshot.Cells, cell => cell.Text == "North");
            Assert.Contains(snapshot.Cells, cell => cell.Text == "South");
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code.StartsWith("ExcelPrintArea", StringComparison.Ordinal));
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code.StartsWith("ExcelPrintArea", StringComparison.Ordinal));
            Assert.True(OfficeImageReader.Identify(png.Bytes).Width > 0);
        }

        [Fact]
        public void ExcelWorksheet_ImageExportExplicitRangeOverridesPrintArea() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            sheet.CellValue(1, 1, "Print");
            sheet.CellValue(3, 3, "Explicit");
            document.SetPrintArea(sheet, "A1:A1", save: false);

            ExcelRangeVisualSnapshot snapshot = sheet.CreateVisualSnapshot(new ExcelWorksheetImageExportOptions {
                Range = "C3:C3",
                UsePrintArea = true
            });

            Assert.Equal("C3:C3", snapshot.Range);
            Assert.Single(snapshot.Cells);
            Assert.Equal("Explicit", snapshot.Cells[0].Text);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code.StartsWith("ExcelPrintArea", StringComparison.Ordinal));
        }

        [Fact]
        public void ExcelWorksheet_ImageExportReportsMissingPrintAreaFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 2, "Value");

            OfficeImageExportResult result = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                UsePrintArea = true
            });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintAreaMissing);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
            Assert.Equal("Report!_xlnm.Print_Area", diagnostic.Source);
        }

        [Fact]
        public void ExcelWorksheet_ImageExportReportsUnsupportedMultiAreaPrintAreaFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, "Used");
                sheet.CellValue(2, 2, "First");
                sheet.CellValue(2, 4, "Second");
                document.Save(false);
            }

            AddMultiAreaPrintArea(filePath);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets[0];
            ExcelRangeVisualSnapshot snapshot = loadedSheet.CreateVisualSnapshot(new ExcelWorksheetImageExportOptions {
                UsePrintArea = true
            });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(snapshot.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Report!_xlnm.Print_Area", diagnostic.Source);
            Assert.Contains(snapshot.Cells, cell => cell.Text == "First");
            Assert.Contains(snapshot.Cells, cell => cell.Text == "Second");
        }

        [Fact]
        public void ExcelWorkbook_ImageExportCanUseWorksheetPrintAreas() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet summary = document.AddWorkSheet("Summary");
            summary.CellValue(1, 1, "Print");
            summary.CellValue(4, 4, "Outside");
            ExcelSheet details = document.AddWorkSheet("Details");
            details.CellValue(1, 1, "No print area");
            document.SetPrintArea(summary, "A1:A1", save: false);
            Assert.Equal("$A$1", summary.GetPrintArea());
            OfficeImageExportResult directSummary = summary.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                UsePrintArea = true
            });
            Assert.Equal("Summary!A1:A1", directSummary.Source);

            IReadOnlyList<OfficeImageExportResult> results = document.ExportImages(OfficeImageExportFormat.Png, new ExcelWorkbookImageExportOptions {
                UseWorksheetPrintAreas = true
            });

            Assert.Equal(2, results.Count);
            Assert.Equal("Summary!A1:A1", results[0].Source);
            Assert.DoesNotContain(results[0].Diagnostics, diagnostic => diagnostic.Code.StartsWith("ExcelPrintArea", StringComparison.Ordinal));
            Assert.Equal("Details!A1:A1", results[1].Source);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintAreaMissing);
            Assert.Equal("Details!_xlnm.Print_Area", diagnostic.Source);
        }

        private static void AddMultiAreaPrintArea(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorkbookPart? workbookPart = spreadsheet.WorkbookPart;
            Assert.NotNull(workbookPart);
            X.Workbook? workbook = workbookPart!.Workbook;
            Assert.NotNull(workbook);
            workbook!.DefinedNames ??= new X.DefinedNames();
            workbook.DefinedNames.Elements<X.DefinedName>()
                .Where(name => string.Equals(name.Name?.Value, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase))
                .ToList()
                .ForEach(name => name.Remove());
            workbook.DefinedNames.Append(new X.DefinedName {
                Name = "_xlnm.Print_Area",
                LocalSheetId = 0U,
                Text = "'Report'!$B$2:$B$2,'Report'!$D$2:$D$2"
            });
            workbook.Save();
        }
    }
}
