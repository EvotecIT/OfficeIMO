using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void ReportWorkflow_TemplateFormulaChartPivot_PreflightsAndExportsPdf() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelReportWorkflow.xlsx");

        try {
            byte[] pdfBytes;
            using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Report")) {
                ExcelSheet sheet = document.Sheets[0];
                sheet.Cell(1, 1, "{{ReportTitle}}");
                sheet.Cell(2, 1, "Region");
                sheet.Cell(2, 2, "Revenue");
                sheet.Cell(2, 3, "Cost");
                sheet.Cell(2, 4, "Margin");
                sheet.Cell(3, 1, "East");
                sheet.Cell(3, 2, 120);
                sheet.Cell(3, 3, 50);
                sheet.CellFormula(3, 4, "B3-C3");
                sheet.Cell(4, 1, "West");
                sheet.Cell(4, 2, 90);
                sheet.Cell(4, 3, 40);
                sheet.CellFormula(4, 4, "B4-C4");

                int replacements = document.ApplyTemplate(new Dictionary<string, object?> {
                    ["ReportTitle"] = "Executive revenue report"
                });
                Assert.Equal(1, replacements);

                Assert.Equal(2, document.RecalculateSupportedFormulas());
                Assert.True(sheet.TryGetCachedFormulaValue(3, 4, out string? eastMargin));
                Assert.Equal("70", eastMargin);

                sheet.AddTable("A2:D4", hasHeader: true, name: "RevenueData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4);
                sheet.AddChartFromRange("A2:D4", row: 6, column: 1, widthPixels: 420, heightPixels: 240,
                    type: ExcelChartType.ColumnClustered, title: "Revenue and Margin");
                sheet.AddPivotTable(
                    sourceRange: "A2:D4",
                    destinationCell: "F2",
                    name: "RevenuePivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Revenue", DataConsolidateFunctionValues.Sum, "Total Revenue") },
                    pivotStyleName: "PivotStyleMedium9");

                ExcelFeatureReport report = document.InspectFeatures();
                Assert.True(report.Can(ExcelPreflightCapability.BindTemplate));
                Assert.True(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.True(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));

                document.Save(false);
                pdfBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                    IncludeSheetHeadings = false,
                    HeaderRowCount = 1,
                    PageSize = new PdfCore.PageSize(560, 520),
                    Margins = PdfCore.PageMargins.Uniform(24)
                });
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(workbookPath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts
                    .Single(part => part.PivotTableParts.Any());
                Assert.Single(worksheetPart.TableDefinitionParts);
                Assert.Single(worksheetPart.PivotTableParts);
                Assert.Single(worksheetPart.DrawingsPart!.ChartParts);
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(pdfBytes));
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Executive revenue report", text);
            Assert.Contains("Revenue and Margin", text);
            Assert.Contains("East", text);
            Assert.Contains("70", text);
        } finally {
            if (File.Exists(workbookPath)) {
                File.Delete(workbookPath);
            }
        }
    }
}
