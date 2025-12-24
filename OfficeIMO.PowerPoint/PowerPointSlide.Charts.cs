using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds a basic clustered column chart with default data.
        /// </summary>
        public PowerPointChart AddChart() {
            return AddChartInternal(null, 0L, 0L, 5486400L, 3200400L);
        }

        /// <summary>
        ///     Adds a clustered column chart using the supplied data.
        /// </summary>
        public PowerPointChart AddChart(PowerPointChartData data, long left = 0L, long top = 0L, long width = 5486400L,
            long height = 3200400L) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            return AddChartInternal(data, left, top, width, height);
        }

        private PowerPointChart AddChartInternal(PowerPointChartData? data, long left, long top, long width, long height) {
            ChartPart chartPart = PowerPointPartFactory.CreatePart<ChartPart>(
                _slidePart,
                contentType: null,
                "/ppt/charts/chart.xml");
            string chartRelId = _slidePart.GetIdOfPart(chartPart);

            // Embed workbook + styles/colors exactly like the template
            ChartStylePart stylePart = PowerPointPartFactory.CreatePart<ChartStylePart>(
                chartPart,
                contentType: null,
                "/ppt/charts/style.xml");
            PowerPointUtils.PopulateChartStyle(stylePart);
            ChartColorStylePart colorStylePart = PowerPointPartFactory.CreatePart<ChartColorStylePart>(
                chartPart,
                contentType: null,
                "/ppt/charts/colors.xml");
            PowerPointUtils.PopulateChartColorStyle(colorStylePart);

            EmbeddedPackagePart embedded = PowerPointPartFactory.CreatePart<EmbeddedPackagePart>(
                chartPart,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx");
            byte[] workbookBytes = data == null ? TemplateChartWorkbookBytes() : PowerPointUtils.BuildChartWorkbook(data);
            using (var ms = new MemoryStream(workbookBytes)) {
                embedded.FeedData(ms);
            }

            string embeddedRelId = chartPart.GetIdOfPart(embedded);
            PowerPointUtils.PopulateChartTemplate(chartPart, embeddedRelId, data);

            string name = GenerateUniqueName("Chart");
            GraphicFrame frame = new(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new Transform(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = chartRelId }) {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                })
            );

            CommonSlideData dataElement = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = dataElement.ShapeTree ??= new ShapeTree();
            tree.AppendChild(frame);
            PowerPointChart chart = new(frame);
            _shapes.Add(chart);
            return chart;
        }

        private static byte[] TemplateChartWorkbookBytes() {
            return PowerPointUtils.GetChartWorkbookTemplateBytes();
        }

        private static byte[] GenerateEmbeddedWorkbookBytes() {
            using MemoryStream ms = new();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new S.Workbook();

                WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
                S.SheetData sheetData = new(
                    new S.Row(
                        new S.Cell { CellValue = new S.CellValue("Category"), DataType = S.CellValues.String },
                        new S.Cell { CellValue = new S.CellValue("Value"), DataType = S.CellValues.String }
                    ),
                    new S.Row(
                        new S.Cell { CellValue = new S.CellValue("A"), DataType = S.CellValues.String },
                        new S.Cell { CellValue = new S.CellValue("4"), DataType = S.CellValues.Number }
                    ),
                    new S.Row(
                        new S.Cell { CellValue = new S.CellValue("B"), DataType = S.CellValues.String },
                        new S.Cell { CellValue = new S.CellValue("5"), DataType = S.CellValues.Number }
                    )
                );
                wsPart.Worksheet = new S.Worksheet(sheetData);

                S.Sheets sheets = new();
                sheets.Append(new S.Sheet { Name = "Sheet1", SheetId = 1U, Id = wbPart.GetIdOfPart(wsPart) });
                wbPart.Workbook.AppendChild(sheets);
                wbPart.Workbook.Save();
            }
            return ms.ToArray();
        }

    }
}
