using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// Demonstrates a template-to-report workflow with formulas, number formats, a chart, a pivot, and an honest PDF preflight decision.
    /// </summary>
    public static class ReportWorkflow
    {
        private static readonly (string Region, string Product, double Revenue, double Cost)[] Sales = {
            ("North", "Platform", 482000, 301000), ("North", "Services", 236000, 171000), ("North", "Support", 118000, 64000),
            ("South", "Platform", 391000, 262000), ("South", "Services", 204000, 158000), ("South", "Support", 97000, 58000),
            ("East", "Platform", 538000, 322000), ("East", "Services", 289000, 201000), ("East", "Support", 141000, 73000),
            ("West", "Platform", 455000, 296000), ("West", "Services", 262000, 189000), ("West", "Support", 126000, 70000)
        };
        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Report workflow");
            string workbookPath = Path.Combine(folderPath, "ExcelReportWorkflow.xlsx");
            string pdfPath = Path.Combine(folderPath, "ExcelReportWorkflow.pdf");
            string previewPath = Path.Combine(folderPath, "ExcelReportWorkflow.png");
            string diagnosticsPath = Path.Combine(folderPath, "ExcelReportWorkflow.preflight.txt");
            using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "Report" }))
            {
                ExcelSheet sheet = document.Sheets[0];
                WriteTitleBand(sheet);
#region Excerpt:operational-excel-report
                // Typed values with number formats; margin and margin % are formulas.
                for (int index = 0; index < Sales.Length; index++)
                {
                    int row = 5 + index;
                    sheet.Cell(row, 1).SetValue(Sales[index].Region);
                    sheet.Cell(row, 2).SetValue(Sales[index].Product);
                    sheet.Cell(row, 3).SetValue(Sales[index].Revenue).SetNumberFormat("#,##0");
                    sheet.Cell(row, 4).SetValue(Sales[index].Cost).SetNumberFormat("#,##0");
                    sheet.Cell(row, 5).SetFormula($"C{row}-D{row}");
                    sheet.Cell(row, 6).SetFormula($"E{row}/C{row}");
                    sheet.Cell(row, 5).SetNumberFormat("#,##0");
                    sheet.Cell(row, 6).SetNumberFormat("0.0%");
                }

                int lastRow = 4 + Sales.Length, totalRow = lastRow + 1;
                sheet.Cell(totalRow, 1).SetValue("Total");
                foreach ((int column, string letter)in new[]
                {
                    (3, "C"),
                    (4, "D"),
                    (5, "E")
                }

                )
                {
                    sheet.Cell(totalRow, column).SetFormula($"SUM({letter}5:{letter}{lastRow})");
                    sheet.Cell(totalRow, column).SetNumberFormat("#,##0");
                }

                sheet.Cell(totalRow, 6).SetFormula($"E{totalRow}/C{totalRow}");
                sheet.Cell(totalRow, 6).SetNumberFormat("0.0%");
                // Table style, a colour scale on margin %, and a frozen header.
                sheet.AddTable($"A4:F{lastRow}", hasHeader: true, name: "RevenueData", style: OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);
                sheet.Range($"F5:F{lastRow}").ConditionalFormat().ColorScale("F8D7DA", "C9ECD3");
                sheet.View.Freeze(topRows: 4);
                // A SUMIF summary by region feeds the chart; the pivot reads the same table.
                string[] regions =
                {
                    "North",
                    "South",
                    "East",
                    "West"
                };
                sheet.Cell(4, 8).SetValue("Region");
                sheet.Cell(4, 9).SetValue("Revenue");
                sheet.Cell(4, 10).SetValue("Margin");
                for (int index = 0; index < regions.Length; index++)
                {
                    int row = 5 + index;
                    sheet.Cell(row, 8).SetValue(regions[index]);
                    sheet.Cell(row, 9).SetFormula($"SUMIF($A$5:$A${lastRow},H{row},$C$5:$C${lastRow})");
                    sheet.Cell(row, 10).SetFormula($"SUMIF($A$5:$A${lastRow},H{row},$E$5:$E${lastRow})");
                }

                sheet.AddChart(OfficeChartType.ColumnClustered, "H4:J8", new OfficeIMO.Excel.ExcelChartOptions { Row = 10, Column = 8, Width = OfficeLength.Pixels(350), Height = OfficeLength.Pixels(235), Title = "Revenue and margin by region" });
                sheet.AddPivotTable($"A4:F{lastRow}", $"A{totalRow + 3}", options: new ExcelPivotTableOptions { Name = "RevenuePivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Revenue", ExcelPivotDataFunction.Sum, "Total Revenue", numberFormat: "#,##0") }, PivotStyleName = "PivotStyleMedium9" });
                // Fill the {{ReportTitle}} placeholder from an object, then calculate every formula.
                document.ApplyTemplate(new { ReportTitle = "Regional revenue report" });
                document.Calculate();
#endregion
                StyleLayout(sheet, totalRow);
                OfficeImageExportResult preview = sheet.Range("A1:K22").ExportImage(OfficeImageExportFormat.Png);
                File.WriteAllBytes(previewPath, preview.Bytes);
                ExcelFeatureReport report = document.InspectFeatures();
                document.Save();
                if (!report.Can(ExcelPreflightCapability.ExportPdfReport))
                {
                    IReadOnlyList<string> diagnostics = report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport);
                    File.WriteAllLines(diagnosticsPath, diagnostics, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                    Console.WriteLine("[!] Excel-to-PDF export is blocked:");
                    foreach (string diagnostic in diagnostics)
                    {
                        Console.WriteLine("  - " + diagnostic);
                    }

                    Console.WriteLine("    Workbook: " + workbookPath);
                    Console.WriteLine("    Preview: " + previewPath);
                    Console.WriteLine("    Diagnostics: " + diagnosticsPath);
                    return;
                }

                document.SaveAsPdf(pdfPath, new ExcelPdfExportOptions { IncludeSheetHeadings = false, HeaderRowCount = 1, PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(560),OfficeIMO.OfficeLength.Points(520)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)) });
                if (openExcel)
                {
                    document.OpenInApplication(workbookPath);
                }
            }
        }

        private static void WriteTitleBand(ExcelSheet sheet) {
            sheet.Range("A1:F1").Merge();
            sheet.Cell(1, 1).SetValue("{{ReportTitle}}")
                .SetFontSize(global::OfficeIMO.OfficeLength.Points(20))
                .SetBold()
                .SetFontColor("FFFFFF");
            sheet.Range("A1:F1").SetFillColor("17365D");
            sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(34)));
            sheet.Range("A2:F2").Merge();
            sheet.Cell(2, 1).SetValue("Revenue, cost and margin by region and product line. Margin columns are formulas.");
            sheet.Cell(2, 1).SetFontColor("5B6B7F");
            string[] headers = { "Region", "Product", "Revenue", "Cost", "Margin", "Margin %" };
            for (int column = 0; column < headers.Length; column++) sheet.Cell(4, column + 1).SetValue(headers[column]);
        }

        private static void StyleLayout(ExcelSheet sheet, int totalRow) {
            sheet.Range(totalRow, 1, totalRow, 6).SetBold().SetFillColor("E8EEF7");
            sheet.Range("I5:J8").SetNumberFormat("#,##0");
            foreach ((int column, double width) in new[] { (1, 12d), (2, 14d), (3, 14d), (4, 14d), (5, 14d), (6, 12d), (7, 3d), (8, 12d), (9, 14d), (10, 14d) }) {
                sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(width));
            }
        }
    }
}