using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates a template-to-report workflow with formulas, charts, pivots, and an honest PDF preflight decision.
    /// </summary>
    public static class ReportWorkflow {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Report workflow");
            string workbookPath = Path.Combine(folderPath, "ExcelReportWorkflow.xlsx");
            string pdfPath = Path.Combine(folderPath, "ExcelReportWorkflow.pdf");
            string previewPath = Path.Combine(folderPath, "ExcelReportWorkflow.png");
            string diagnosticsPath = Path.Combine(folderPath, "ExcelReportWorkflow.preflight.txt");

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

                OfficeImageExportResult preview = sheet.Range("A1:J20").ExportImage(OfficeImageExportFormat.Png);
                File.WriteAllBytes(previewPath, preview.Bytes);

                ExcelFeatureReport report = document.InspectFeatures();
                document.Save();
                if (!report.Can(ExcelPreflightCapability.ExportPdfReport)) {
                    IReadOnlyList<string> diagnostics = report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport);
                    File.WriteAllLines(diagnosticsPath, diagnostics, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                    Console.WriteLine("[!] Excel-to-PDF export is blocked:");
                    foreach (string diagnostic in diagnostics) {
                        Console.WriteLine("  - " + diagnostic);
                    }
                    Console.WriteLine("    Workbook: " + workbookPath);
                    Console.WriteLine("    Preview: " + previewPath);
                    Console.WriteLine("    Diagnostics: " + diagnosticsPath);
                    return;
                }

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
