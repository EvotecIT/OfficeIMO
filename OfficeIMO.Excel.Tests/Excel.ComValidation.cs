using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeColor = OfficeIMO.Drawing.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static readonly TimeSpan ExcelComOpenTimeout = TimeSpan.FromMinutes(2);

#if NET5_0_OR_GREATER
        [SupportedOSPlatformGuard("windows")]
#endif
        private static bool IsWindowsPlatform() =>
            RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static bool IsExcelComAvailable() =>
            Type.GetTypeFromProgID("Excel.Application") != null;

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void AssertWorkbookOpensViaExcelComWhenAvailable(string path, string failureMessage) =>
            AssertWorkbooksOpenViaExcelComWhenAvailable(new[] { path }, failureMessage);

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void AssertWorkbooksOpenViaExcelComWhenAvailable(IEnumerable<string> paths, string failureMessage) {
            if (!IsExcelComAvailable()) {
                return;
            }

            List<string> failures = new();
            var thread = new Thread(() => OpenWorkbooksViaExcelCom(paths.ToList(), failures));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(ExcelComOpenTimeout)) {
                failures.Add($"Excel COM smoke test timed out after {ExcelComOpenTimeout.TotalSeconds:0} seconds.");
            }

            Assert.True(failures.Count == 0, failureMessage + Environment.NewLine + string.Join(Environment.NewLine, failures));
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void OpenWorkbooksViaExcelCom(IReadOnlyList<string> paths, List<string> failures) {
            object? excel = null;
            object? workbooks = null;

            try {
                var excelType = Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException("Excel COM automation is not available.");
                excel = Activator.CreateInstance(excelType)
                    ?? throw new InvalidOperationException("Failed to create Excel COM automation instance.");

                excelType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, excel, new object[] { false });
                excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, excel, new object[] { false });
                workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, excel, null);

                foreach (string path in paths) {
                    object? workbook = null;
                    try {
                        workbook = workbooks!.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks,
                            new object[] { path, 0, true });
                    } catch (Exception ex) when (ex is COMException or InvalidOperationException or MissingMethodException or TargetInvocationException) {
                        failures.Add($"{Path.GetFileName(path)}: {DescribeExcelComFailure(ex)}");
                    } finally {
                        try {
                            workbook?.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                        } catch (Exception ex) when (ex is COMException or MissingMethodException or TargetInvocationException) {
                            failures.Add($"{Path.GetFileName(path)} close: {DescribeExcelComFailure(ex)}");
                        }

                        if (workbook != null && Marshal.IsComObject(workbook)) {
                            Marshal.FinalReleaseComObject(workbook);
                        }
                    }
                }
            } catch (Exception ex) when (ex is COMException or InvalidOperationException or MissingMethodException or TargetInvocationException) {
                failures.Add(DescribeExcelComFailure(ex));
            } finally {
                try {
                    excel?.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excel, null);
                } catch (Exception ex) when (ex is COMException or MissingMethodException or TargetInvocationException) {
                    failures.Add("Excel quit: " + DescribeExcelComFailure(ex));
                }

                if (workbooks != null && Marshal.IsComObject(workbooks)) {
                    Marshal.FinalReleaseComObject(workbooks);
                }
                if (excel != null && Marshal.IsComObject(excel)) {
                    Marshal.FinalReleaseComObject(excel);
                }
            }
        }

        private static string DescribeExcelComFailure(Exception ex) {
            Exception actual = ex is TargetInvocationException { InnerException: not null } target
                ? target.InnerException!
                : ex;
            return actual.GetType().Name + ": " + actual.Message;
        }

        [Fact]
        public void Test_ExcelWorkbooks_OpenInDesktopExcelWhenAvailable() {
            if (!IsWindowsPlatform()) {
                return;
            }

            string directory = Path.Combine(_directoryWithFiles, "DesktopExcelSmoke");
            Directory.CreateDirectory(directory);

            var files = new[] {
                CreateDesktopExcelValuesWorkbook(Path.Combine(directory, "Values.Formulas.xlsx")),
                CreateDesktopExcelTableWorkbook(Path.Combine(directory, "Table.Filter.Freeze.xlsx")),
                CreateDesktopExcelFormattingWorkbook(Path.Combine(directory, "Formatting.Validation.xlsx")),
                CreateDesktopExcelChartWorkbook(Path.Combine(directory, "Charts.ComboScatter.xlsx")),
                CreateDesktopExcelRecipeChartWorkbook(Path.Combine(directory, "Charts.Recipes.xlsx")),
                CreateDesktopExcelModernChartWorkbook(Path.Combine(directory, "Charts.ModernCompatible.xlsx")),
                CreateDesktopExcelPivotInteractionWorkbook(Path.Combine(directory, "Pivot.Interactions.xlsx"))
            };

            AssertWorkbooksOpenViaExcelComWhenAvailable(files,
                "One or more generated Excel workbooks required repair or failed to open in desktop Excel.");
        }

        private static string CreateDesktopExcelValuesWorkbook(string filePath) {
            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Values");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, "Alpha");
            sheet.CellValue(2, 2, 10d);
            sheet.CellValue(3, 1, "Beta");
            sheet.CellValue(3, 2, 20.5d);
            sheet.CellFormula(4, 2, "SUM(B2:B3)");
            sheet.Range("A1:B1").HeaderStyle();
            document.Save();
            return filePath;
        }

        private static string CreateDesktopExcelTableWorkbook(string filePath) {
            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Score");
            sheet.CellValue(2, 1, "North");
            sheet.CellValue(2, 2, 80d);
            sheet.CellValue(3, 1, "South");
            sheet.CellValue(3, 2, 95d);
            sheet.AddTable("A1:B3", true, "ScoresTable", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
            sheet.AddAutoFilter("A1:B3", new Dictionary<uint, IEnumerable<string>> {
                { 0, new[] { "North", "South" } }
            });
            sheet.Freeze(topRows: 1, leftCols: 1);
            document.Save();
            return filePath;
        }

        private static string CreateDesktopExcelFormattingWorkbook(string filePath) {
            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Formatting");
            sheet.CellValue(1, 1, "Score");
            sheet.CellValue(2, 1, 42d);
            sheet.CellValue(3, 1, 87d);
            sheet.AddConditionalColorScale("A2:A3", OfficeColor.Red, OfficeColor.Lime);
            sheet.AddConditionalDataBar("A2:A3", OfficeColor.Blue);
            sheet.AddConditionalRule("A2:A3", ConditionalFormattingOperatorValues.GreaterThan, "50");
            document.Save();
            return filePath;
        }

        private static string CreateDesktopExcelChartWorkbook(string filePath) {
            using var document = ExcelDocument.Create(filePath);
            document.DefaultChartStylePreset = ExcelChartStylePreset.Default;
            ExcelSheet sheet = document.AddWorksheet("Summary");

            var data = new ExcelChartData(
                new[] { "Q1", "Q2", "Q3", "Q4" },
                new[] {
                    new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }, ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary),
                    new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d, 35d }, ExcelChartType.Line, ExcelChartAxisGroup.Secondary)
                });

            ExcelChart chart = sheet.AddChart(data, row: 2, column: 6, widthPixels: 640, heightPixels: 360,
                type: ExcelChartType.ColumnClustered, title: "Sales vs Trend");
            chart.SetSeriesDataLabels(1, showValue: true, position: DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues.Top)
                 .SetSeriesDataLabelForPoint(1, 2, showValue: true, position: DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues.OutsideEnd);

            document.Save();
            return filePath;
        }

        private static string CreateDesktopExcelRecipeChartWorkbook(string filePath) {
            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Revenue");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 10d);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 16d);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 13d);
            sheet.AddTable("A1:B4", hasHeader: true, name: "RevenueData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium9);

            sheet.AddRevenueTrendChart("A1:B4", row: 1, column: 5);
            sheet.AddTopNBarChart("A1:B4", row: 18, column: 5, title: "Top Revenue");
            sheet.AddVarianceColumnChart("A1:B4", row: 35, column: 5, title: "Revenue Variance");
            sheet.AddKpiScorecardChart("A1:B4", row: 52, column: 5, title: "Revenue KPI");
            sheet.AddContributionChart("A1:B4", row: 69, column: 5, title: "Revenue Mix");
            sheet.ChartFromTable("RevenueData")
                .StatusBreakdown("Revenue Mix")
                .At(86, 5);
            sheet.ChartFromTable("RevenueData")
                .VarianceWaterfall("Revenue Bridge")
                .At(103, 5);

            document.Save();
            return filePath;
        }

        private static string CreateDesktopExcelModernChartWorkbook(string filePath) {
            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            sheet.AddHistogramChart(new[] { 1d, 2d, 2d, 3d, 5d, 8d }, row: 1, column: 1, binCount: 3);
            sheet.AddParetoChart(new[] { "Returns", "Delays", "Damage" }, new[] { 8d, 13d, 3d }, row: 1, column: 10);
            sheet.AddFunnelChart(new[] { "Leads", "Qualified", "Won" }, new[] { 120d, 70d, 25d }, row: 20, column: 1);
            sheet.AddWaterfallChart(new[] { "Opening", "Growth", "Cost" }, new[] { 100d, 35d, -20d }, row: 20, column: 10);
            document.Save();
            return filePath;
        }

        private static string CreateDesktopExcelPivotInteractionWorkbook(string filePath) {
            using var document = ExcelDocument.Create(filePath);
            ExcelSheet source = document.AddWorksheet("Source");
            WriteDesktopPivotSource(source, 3);
            source.AddPivotTable(
                sourceRange: "A1:C4",
                destinationCell: "E2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales") });

            ExcelSheet expanded = document.AddWorksheet("Expanded");
            WriteDesktopPivotSource(expanded, 4);
            source.UpdatePivotTableSource("SalesPivot", expanded, "$A$1:$C$5");
            document.AddPivotSlicerCache("SalesPivot", "Region");
            document.AddPivotTimelineCache("SalesPivot", "Date");
            document.Save();
            return filePath;
        }

        private static void WriteDesktopPivotSource(ExcelSheet sheet, int dataRows) {
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Date");
            sheet.CellValue(1, 3, "Sales");
            for (int index = 0; index < dataRows; index++) {
                sheet.CellValue(index + 2, 1, index % 2 == 0 ? "East" : "West");
                sheet.CellValue(index + 2, 2, new DateTime(2026, index + 1, 1));
                sheet.CellValue(index + 2, 3, (index + 1) * 100d);
            }
        }
    }
}
