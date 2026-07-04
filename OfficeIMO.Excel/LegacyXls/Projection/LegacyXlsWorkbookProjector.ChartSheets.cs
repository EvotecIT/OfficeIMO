using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SpreadsheetDrawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static partial class LegacyXlsWorkbookProjector {
        private const string ChartGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        private static void ProjectChartSheets(LegacyXlsWorkbook workbook, ExcelDocument document) {
            if (workbook.ChartSheets.Count == 0) {
                return;
            }

            WorkbookPart workbookPart = document.WorkbookPartRoot;
            Workbook workbookRoot = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook part is missing a workbook root.");
            Sheets sheets = workbookRoot.Sheets ?? workbookRoot.AppendChild(new Sheets());
            uint nextSheetId = GetNextSheetId(sheets);

            foreach (LegacyXlsChartSheet chartSheet in workbook.ChartSheets) {
                ChartsheetPart chartsheetPart = workbookPart.AddNewPart<ChartsheetPart>();
                DrawingsPart drawingsPart = chartsheetPart.AddNewPart<DrawingsPart>();
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = CreateChartSheetChartSpace(chartSheet);
                chartPart.ChartSpace.Save();

                string chartRelationshipId = drawingsPart.GetIdOfPart(chartPart);
                drawingsPart.WorksheetDrawing = CreateChartSheetDrawing(chartSheet, chartRelationshipId);
                drawingsPart.WorksheetDrawing.Save();

                string drawingRelationshipId = chartsheetPart.GetIdOfPart(drawingsPart);
                chartsheetPart.Chartsheet = CreateChartsheet(chartSheet, drawingRelationshipId);
                chartsheetPart.Chartsheet.Save();

                var sheet = new Sheet {
                    Id = workbookPart.GetIdOfPart(chartsheetPart),
                    SheetId = nextSheetId++,
                    Name = chartSheet.Name
                };

                SheetStateValues? state = ToSheetState(chartSheet.VisibilityKind);
                if (state.HasValue) {
                    sheet.State = state.Value;
                }

                sheets.Append(sheet);
            }

            workbookPart.Workbook.Save();
        }

        private static uint GetNextSheetId(Sheets sheets) {
            uint currentMax = 0U;
            foreach (Sheet sheet in sheets.Elements<Sheet>()) {
                uint value = sheet.SheetId?.Value ?? 0U;
                if (value > currentMax) {
                    currentMax = value;
                }
            }

            return currentMax + 1U;
        }

        private static SheetStateValues? ToSheetState(LegacyXlsSheetVisibility? visibility) {
            switch (visibility) {
                case LegacyXlsSheetVisibility.Hidden:
                    return SheetStateValues.Hidden;
                case LegacyXlsSheetVisibility.VeryHidden:
                    return SheetStateValues.VeryHidden;
                default:
                    return null;
            }
        }

        private static Chartsheet CreateChartsheet(LegacyXlsChartSheet chartSheet, string chartRelationshipId) {
            var chartsheet = new Chartsheet(
                new SheetViews(
                    new SheetView {
                        WorkbookViewId = 0U
                    }),
                new PageMargins {
                    Left = 0.7D,
                    Right = 0.7D,
                    Top = 0.75D,
                    Bottom = 0.75D,
                    Header = 0.3D,
                    Footer = 0.3D
                },
                new SpreadsheetDrawing {
                    Id = chartRelationshipId
                });

            return chartsheet;
        }

        private static Xdr.WorksheetDrawing CreateChartSheetDrawing(LegacyXlsChartSheet chartSheet, string chartRelationshipId) {
            return new Xdr.WorksheetDrawing(
                new Xdr.AbsoluteAnchor(
                    new Xdr.Position {
                        X = 0L,
                        Y = 0L
                    },
                    new Xdr.Extent {
                        Cx = 9144000L,
                        Cy = 6858000L
                    },
                    new Xdr.GraphicFrame(
                        new Xdr.NonVisualGraphicFrameProperties(
                            new Xdr.NonVisualDrawingProperties {
                                Id = 2U,
                                Name = chartSheet.Name
                            },
                            new Xdr.NonVisualGraphicFrameDrawingProperties()),
                        new Xdr.Transform(
                            new A.Offset {
                                X = 0L,
                                Y = 0L
                            },
                            new A.Extents {
                                Cx = 9144000L,
                                Cy = 6858000L
                            }),
                        new A.Graphic(
                            new A.GraphicData(
                                new C.ChartReference {
                                    Id = chartRelationshipId
                                }) {
                                Uri = ChartGraphicDataUri
                            })),
                    new Xdr.ClientData()));
        }

        private static C.ChartSpace CreateChartSheetChartSpace(LegacyXlsChartSheet chartSheet) {
            C.PlotArea plotArea = CreateChartSheetPlotArea(chartSheet);
            return new C.ChartSpace(
                new C.EditingLanguage { Val = "en-US" },
                new C.RoundedCorners { Val = false },
                new C.Chart(
                    CreateChartTitle(chartSheet.Name),
                    new C.AutoTitleDeleted { Val = false },
                    plotArea,
                    new C.PlotVisibleOnly { Val = true },
                    new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap }),
                new C.PrintSettings(
                    new C.HeaderFooter(),
                    new C.PageMargins {
                        Left = 0.7D,
                        Right = 0.7D,
                        Top = 0.75D,
                        Bottom = 0.75D,
                        Header = 0.3D,
                        Footer = 0.3D
                    },
                    new C.PageSetup()));
        }

        private static C.Title CreateChartTitle(string title) {
            return new C.Title(
                new C.ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text(title))))),
                new C.Layout(),
                new C.Overlay { Val = false });
        }

        private static C.PlotArea CreateChartSheetPlotArea(LegacyXlsChartSheet chartSheet) {
            var plotArea = new C.PlotArea(new C.Layout());
            string chartType = chartSheet.ChartRecordsByChartType.Keys.FirstOrDefault() ?? "Bar";
            switch (chartType) {
                case "Line":
                    plotArea.Append(CreateEmptyLineChart());
                    AppendCategoryAndValueAxes(plotArea);
                    break;
                case "Pie":
                case "BarOfPieOrPieOfPie":
                    plotArea.Append(CreateEmptyPieChart());
                    break;
                case "Area":
                    plotArea.Append(CreateEmptyAreaChart());
                    AppendCategoryAndValueAxes(plotArea);
                    break;
                case "Scatter":
                    plotArea.Append(CreateEmptyScatterChart());
                    AppendValueAxes(plotArea);
                    break;
                default:
                    plotArea.Append(CreateEmptyBarChart());
                    AppendCategoryAndValueAxes(plotArea);
                    break;
            }

            return plotArea;
        }

        private static C.BarChart CreateEmptyBarChart() {
            return new C.BarChart(
                new C.BarDirection { Val = C.BarDirectionValues.Column },
                new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                new C.VaryColors { Val = false },
                new C.AxisId { Val = 48650112U },
                new C.AxisId { Val = 48672768U });
        }

        private static C.LineChart CreateEmptyLineChart() {
            return new C.LineChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false },
                new C.AxisId { Val = 48650112U },
                new C.AxisId { Val = 48672768U });
        }

        private static C.AreaChart CreateEmptyAreaChart() {
            return new C.AreaChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false },
                new C.AxisId { Val = 48650112U },
                new C.AxisId { Val = 48672768U });
        }

        private static C.PieChart CreateEmptyPieChart() {
            return new C.PieChart(
                new C.VaryColors { Val = true },
                new C.FirstSliceAngle { Val = (ushort)0 });
        }

        private static C.ScatterChart CreateEmptyScatterChart() {
            return new C.ScatterChart(
                new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                new C.VaryColors { Val = false },
                new C.AxisId { Val = 48650112U },
                new C.AxisId { Val = 48672768U });
        }

        private static void AppendCategoryAndValueAxes(C.PlotArea plotArea) {
            plotArea.Append(
                new C.CategoryAxis(
                    new C.AxisId { Val = 48650112U },
                    new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                    new C.Delete { Val = false },
                    new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                    new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                    new C.CrossingAxis { Val = 48672768U },
                    new C.Crosses { Val = C.CrossesValues.AutoZero },
                    new C.AutoLabeled { Val = true },
                    new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
                    new C.LabelOffset { Val = (ushort)100 }),
                CreateValueAxis(48672768U, 48650112U, C.AxisPositionValues.Left));
        }

        private static void AppendValueAxes(C.PlotArea plotArea) {
            plotArea.Append(
                CreateValueAxis(48650112U, 48672768U, C.AxisPositionValues.Bottom),
                CreateValueAxis(48672768U, 48650112U, C.AxisPositionValues.Left));
        }

        private static C.ValueAxis CreateValueAxis(uint axisId, uint crossingAxisId, C.AxisPositionValues position) {
            return new C.ValueAxis(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = position },
                new C.MajorGridlines(),
                new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = crossingAxisId },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween { Val = C.CrossBetweenValues.Between });
        }
    }
}
