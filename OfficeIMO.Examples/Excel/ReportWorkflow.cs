using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates a template-to-report workflow with formulas, charts, pivots, preflight, and PDF export.
    /// </summary>
    public static class ReportWorkflow {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Report workflow");
            string workbookPath = Path.Combine(folderPath, "ExcelReportWorkflow.xlsx");
            string pdfPath = Path.Combine(folderPath, "ExcelReportWorkflow.pdf");

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

                document.ApplyTemplate(new {
                    ReportTitle = "Executive revenue report"
                });
                document.Calculate();

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
                if (!report.Can(ExcelPreflightCapability.ExportPdfReport)) {
                    Console.WriteLine("[!] Excel-to-PDF export is blocked:");
                    foreach (string diagnostic in report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport)) {
                        Console.WriteLine("  - " + diagnostic);
                    }
                    return;
                }

                document.Save();
                document.SaveAsPdf(pdfPath, new ExcelPdfSaveOptions {
                    IncludeSheetHeadings = false,
                    HeaderRowCount = 1,
                    PageSize = new PdfCore.PageSize(560, 520),
                    Margins = PdfCore.PageMargins.Uniform(24)
                });

                if (openExcel) {
                    document.OpenInApplication(workbookPath);
                }
            }
        }
    }
}
