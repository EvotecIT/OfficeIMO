using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

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
            string chartPartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/charts",
                "chart",
                ".xml",
                allowBaseWithoutIndex: false);
            ChartPart chartPart = PowerPointPartFactory.CreatePart<ChartPart>(
                _slidePart,
                contentType: null,
                chartPartUri);
            string chartRelId = _slidePart.GetIdOfPart(chartPart);

            // Embed workbook + styles/colors exactly like the template
            string stylePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/charts",
                "style",
                ".xml",
                allowBaseWithoutIndex: false);
            ChartStylePart stylePart = PowerPointPartFactory.CreatePart<ChartStylePart>(
                chartPart,
                contentType: null,
                stylePartUri);
            PowerPointUtils.PopulateChartStyle(stylePart);
            string colorStylePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/charts",
                "colors",
                ".xml",
                allowBaseWithoutIndex: false);
            ChartColorStylePart colorStylePart = PowerPointPartFactory.CreatePart<ChartColorStylePart>(
                chartPart,
                contentType: null,
                colorStylePartUri);
            PowerPointUtils.PopulateChartColorStyle(colorStylePart);

            string embeddedPartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/embeddings",
                "Microsoft_Excel_Worksheet",
                ".xlsx",
                allowBaseWithoutIndex: false);
            EmbeddedPackagePart embedded = PowerPointPartFactory.CreatePart<EmbeddedPackagePart>(
                chartPart,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                embeddedPartUri);
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

    }
}
