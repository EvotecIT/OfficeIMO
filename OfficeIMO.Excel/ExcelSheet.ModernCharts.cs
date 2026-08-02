using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string ChartExGraphicDataUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";

        /// <summary>Enumerates native and imported ChartEx drawings on this worksheet.</summary>
        public IEnumerable<ExcelModernChart> ModernCharts {
            get {
                DrawingsPart? drawingsPart = _worksheetPart.DrawingsPart;
                if (drawingsPart?.WorksheetDrawing == null) return Enumerable.Empty<ExcelModernChart>();
                return drawingsPart.WorksheetDrawing.Descendants<Xdr.GraphicFrame>()
                    .Where(frame => string.Equals(
                        frame.Graphic?.GraphicData?.Uri?.Value,
                        ChartExGraphicDataUri,
                        StringComparison.Ordinal))
                    .Select(frame => {
                        ExtendedChartPart? part = TryGetModernChartPart(frame, drawingsPart);
                        return new ExcelModernChart(
                            frame,
                            drawingsPart,
                            this,
                            part == null ? null : TryExtractModernChartDataRange(part));
                    })
                    .ToArray();
            }
        }

        /// <summary>Returns a native or imported ChartEx drawing by non-visual name.</summary>
        public ExcelModernChart? GetModernChart(string name) {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return ModernCharts.FirstOrDefault(chart => string.Equals(chart.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Adds a native ChartEx chart backed by OfficeIMO's shared hidden chart-data owner.</summary>
        public ExcelModernChart AddModernChart(
            ExcelChartData data,
            int row,
            int column,
            ExcelModernChartType chartType,
            string? title = null,
            int widthPixels = 640,
            int heightPixels = 360) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            ValidateModernChartPlacement(row, column, widthPixels, heightPixels);
            string layout = GetModernChartLayout(chartType);
            ValidateModernChartData(data);
            ValidateModernChartText(title, nameof(title));

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateDuringDirtyMarks();
            ExcelChartDataRange? range = null;
            Xdr.GraphicFrame? frame = null;
            DrawingsPart? drawingsPart = null;
            WriteLock(() => {
                DocumentFormat.OpenXml.Spreadsheet.Drawing? drawing = WorksheetRoot
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
                if (drawing != null) {
                    string drawingRelationshipId = drawing.Id?.Value ?? string.Empty;
                    if (drawingRelationshipId.Length == 0
                        || TryGetPartById(_worksheetPart, drawingRelationshipId) is not DrawingsPart existingDrawingsPart) {
                        throw new InvalidOperationException("Worksheet drawing relationship is missing or invalid.");
                    }
                    drawingsPart = existingDrawingsPart;
                }

                ExcelSheet dataSheet = _excelDocument.GetOrCreateChartDataSheet();
                int startRow = _excelDocument.ReserveChartDataStartRow(
                    dataSheet,
                    GetChartDataPointCount(data) + 1);
                range = dataSheet.WriteChartData(data, startRow, 1);
                if (drawing == null) {
                    drawingsPart = _worksheetPart.AddNewPart<DrawingsPart>();
                    drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                    WorksheetRoot.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing {
                        Id = _worksheetPart.GetIdOfPart(drawingsPart)
                    });
                } else {
                    drawingsPart!.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
                }

                ExtendedChartPart chartPart = drawingsPart!.AddNewPart<ExtendedChartPart>();
                chartPart.ChartSpace = BuildModernChartSpace(data, range!, layout, title);
                chartPart.ChartSpace.Save();
                string relationshipId = drawingsPart.GetIdOfPart(chartPart);
                long width = PxToEmu(widthPixels);
                long height = PxToEmu(heightPixels);
                UInt32Value id = NextDrawingId(drawingsPart);
                string name = "Modern Chart " + id.Value.ToString(CultureInfo.InvariantCulture);
                frame = new Xdr.GraphicFrame(
                    new Xdr.NonVisualGraphicFrameProperties(
                        new Xdr.NonVisualDrawingProperties { Id = id, Name = name },
                        new Xdr.NonVisualGraphicFrameDrawingProperties(new GraphicFrameLocks { NoChangeAspect = true })),
                    new Xdr.Transform(
                        new Offset { X = 0, Y = 0 },
                        new Extents { Cx = width, Cy = height }),
                    new Graphic(
                        new GraphicData(
                            new Cx.RelId { Id = relationshipId }) { Uri = ChartExGraphicDataUri }));
                drawingsPart.WorksheetDrawing!.Append(new Xdr.OneCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture)),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture)),
                        new Xdr.RowOffset("0")),
                    new Xdr.Extent { Cx = width, Cy = height },
                    frame,
                    new Xdr.ClientData()));
                drawingsPart.WorksheetDrawing.Save();
                MarkRequiresSavePreparation();
            });
            return new ExcelModernChart(frame!, drawingsPart!, this, range!);
        }

        internal static void ValidateModernChartPlacement(int row, int column, int widthPixels, int heightPixels) {
            if (row < 1 || row > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1 || column > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
            if (widthPixels < 1) throw new ArgumentOutOfRangeException(nameof(widthPixels));
            if (heightPixels < 1) throw new ArgumentOutOfRangeException(nameof(heightPixels));
        }

        internal static void ValidateModernChartData(ExcelChartData data) {
            if (data.Series.Count == 0 || data.Categories.Count == 0) {
                throw new ArgumentException("Modern charts require at least one category and one series.", nameof(data));
            }
            if (data.Series.Count >= A1.MaxColumns) {
                throw new ArgumentException(
                    "Modern chart data must fit one category column and all series within the worksheet column limit.",
                    nameof(data));
            }
            if (data.Series.Any(series => series.Values.Count != data.Categories.Count)) {
                throw new ArgumentException("Modern chart series must match the category count.", nameof(data));
            }
            if (data.Series.Any(series => series.Values.Any(value => double.IsNaN(value) || double.IsInfinity(value)))) {
                throw new ArgumentException("Modern chart series values must be finite numbers.", nameof(data));
            }
            if (data.Categories.Any(category => !IsValidModernChartText(category))
                || data.Series.Any(series => !IsValidModernChartText(series.Name))) {
                throw new ArgumentException(
                    "Modern chart categories and series names must contain valid Excel XML text.",
                    nameof(data));
            }
        }

        internal static void ValidateModernChartText(string? text, string parameterName) {
            if (!IsValidModernChartText(text)) {
                throw new ArgumentException("Modern chart text must contain valid Excel XML text.", parameterName);
            }
        }

        private static bool IsValidModernChartText(string? text) {
            string value = text ?? string.Empty;
            if (value.Length > 32_767) return false;
            try {
                XmlConvert.VerifyXmlChars(value);
                return true;
            } catch (XmlException) {
                return false;
            }
        }

        internal static string GetModernChartLayout(ExcelModernChartType chartType) {
            return chartType switch {
                ExcelModernChartType.Funnel => "funnel",
                ExcelModernChartType.Waterfall => "waterfall",
                ExcelModernChartType.BoxWhisker => "boxWhisker",
                ExcelModernChartType.Treemap => "treemap",
                ExcelModernChartType.Sunburst => "sunburst",
                _ => throw new ArgumentOutOfRangeException(nameof(chartType), "Unsupported is inspection-only and cannot be authored.")
            };
        }

        internal static ExcelModernChartType ParseModernChartType(string? layout) {
            return layout switch {
                "funnel" => ExcelModernChartType.Funnel,
                "waterfall" => ExcelModernChartType.Waterfall,
                "boxWhisker" => ExcelModernChartType.BoxWhisker,
                "treemap" => ExcelModernChartType.Treemap,
                "sunburst" => ExcelModernChartType.Sunburst,
                _ => ExcelModernChartType.Unsupported
            };
        }

        private static ExtendedChartPart? TryGetModernChartPart(Xdr.GraphicFrame frame, DrawingsPart drawingsPart) {
            OpenXmlElement? reference = frame.Graphic?.GraphicData?.ChildElements.FirstOrDefault(element =>
                string.Equals(element.LocalName, "chart", StringComparison.Ordinal)
                && string.Equals(element.NamespaceUri, ChartExGraphicDataUri, StringComparison.Ordinal));
            string relationshipId = reference?.GetAttribute(
                "id",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value ?? string.Empty;
            if (relationshipId.Length == 0) return null;
            return drawingsPart.Parts.FirstOrDefault(pair =>
                string.Equals(pair.RelationshipId, relationshipId, StringComparison.Ordinal)).OpenXmlPart as ExtendedChartPart;
        }

        private ExcelChartDataRange? TryExtractModernChartDataRange(ExtendedChartPart part) {
            OpenXmlElement? root = part.ChartSpace;
            OpenXmlElement[] data = root?.Descendants()
                .Where(element => element.LocalName == "data")
                .ToArray() ?? Array.Empty<OpenXmlElement>();
            if (data.Length == 0) return null;

            string? categoryFormula = data[0].Descendants()
                .FirstOrDefault(element => element.LocalName == "strDim"
                    && element.GetAttribute("type", string.Empty).Value == "cat")?
                .Descendants()
                .FirstOrDefault(element => element.LocalName == "f")?
                .InnerText;
            if (!ExcelChartUtils.TryParseSheetQualifiedRange(categoryFormula, out string sheetName, out string categoryRange)
                || !_excelDocument.IsOwnedChartDataSheet(sheetName)
                || !A1.TryParseRange(categoryRange, out int firstRow, out int categoryColumn, out int lastRow, out int lastCategoryColumn)
                || categoryColumn != lastCategoryColumn) return null;

            for (int index = 0; index < data.Length; index++) {
                string? valueFormula = data[index].Descendants()
                    .FirstOrDefault(element => element.LocalName == "numDim"
                        && element.GetAttribute("type", string.Empty).Value == "val")?
                    .Descendants()
                    .FirstOrDefault(element => element.LocalName == "f")?
                    .InnerText;
                if (!ExcelChartUtils.TryParseSheetQualifiedRange(valueFormula, out string valueSheet, out string valueRange)
                    || !string.Equals(sheetName, valueSheet, StringComparison.OrdinalIgnoreCase)
                    || !A1.TryParseRange(valueRange, out int valueFirstRow, out int valueColumn, out int valueLastRow, out int valueLastColumn)
                    || valueFirstRow != firstRow
                    || valueLastRow != lastRow
                    || valueColumn != categoryColumn + index + 1
                    || valueLastColumn != valueColumn) return null;
            }

            bool hasHeaderRow = firstRow > 1;
            if (hasHeaderRow) {
                ExcelSheet dataSheet = _excelDocument[sheetName];
                Cx.Series[] series = root?.Descendants<Cx.Series>().ToArray() ?? Array.Empty<Cx.Series>();
                hasHeaderRow = series.Length == data.Length;
                for (int index = 0; hasHeaderRow && index < series.Length; index++) {
                    string expectedName = series[index].Descendants<Cx.VXsdstring>().FirstOrDefault()?.Text ?? string.Empty;
                    hasHeaderRow = dataSheet.TryGetCellText(firstRow - 1, categoryColumn + index + 1, out string actualName)
                        && string.Equals(expectedName, actualName, StringComparison.Ordinal);
                }
            }

            return new ExcelChartDataRange(
                sheetName,
                hasHeaderRow ? firstRow - 1 : firstRow,
                categoryColumn,
                lastRow - firstRow + 1,
                data.Length,
                hasHeaderRow);
        }

        internal static Cx.ChartSpace BuildModernChartSpace(
            ExcelChartData data,
            ExcelChartDataRange range,
            string layout,
            string? title) {
            XNamespace cx = ChartExGraphicDataUri;
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var chart = new XElement(cx + "chart");
            if (!string.IsNullOrWhiteSpace(title)) {
                chart.Add(new XElement(cx + "title",
                    new XAttribute("pos", "t"),
                    new XElement(cx + "tx",
                        new XElement(cx + "txData",
                            new XElement(cx + "v", title!.Trim())))));
            }
            var region = new XElement(cx + "plotAreaRegion");
            var chartData = new XElement(cx + "chartData");
            string categoriesFormula = ExcelChartUtils.BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            for (int seriesIndex = 0; seriesIndex < data.Series.Count; seriesIndex++) {
                ExcelChartSeries series = data.Series[seriesIndex];
                region.Add(new XElement(cx + "series",
                    new XAttribute("layoutId", layout),
                    new XAttribute("ownerIdx", seriesIndex.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("uniqueId", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"),
                    new XElement(cx + "tx",
                        new XElement(cx + "txData", new XElement(cx + "v", series.Name))),
                    new XElement(cx + "dataId", new XAttribute("val", seriesIndex.ToString(CultureInfo.InvariantCulture)))));

                var categoryLevel = new XElement(cx + "lvl", new XAttribute("ptCount", data.Categories.Count));
                for (int index = 0; index < data.Categories.Count; index++) {
                    categoryLevel.Add(new XElement(cx + "pt",
                        new XAttribute("idx", index),
                        data.Categories[index] ?? string.Empty));
                }
                var valueLevel = new XElement(cx + "lvl",
                    new XAttribute("ptCount", series.Values.Count),
                    new XAttribute("formatCode", "General"));
                for (int index = 0; index < series.Values.Count; index++) {
                    valueLevel.Add(new XElement(cx + "pt",
                        new XAttribute("idx", index),
                        series.Values[index].ToString("R", CultureInfo.InvariantCulture)));
                }
                chartData.Add(new XElement(cx + "data",
                    new XAttribute("id", seriesIndex),
                    new XElement(cx + "strDim",
                        new XAttribute("type", "cat"),
                        new XElement(cx + "f", categoriesFormula),
                        categoryLevel),
                    new XElement(cx + "numDim",
                        new XAttribute("type", "val"),
                        new XElement(cx + "f", ExcelChartUtils.BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex))),
                        valueLevel)));
            }
            chart.Add(new XElement(cx + "plotArea", region));
            var root = new XElement(cx + "chartSpace",
                new XAttribute(XNamespace.Xmlns + "cx", cx),
                new XAttribute(XNamespace.Xmlns + "a", a),
                new XAttribute(XNamespace.Xmlns + "r", r),
                new XAttribute("version", "1"),
                chart,
                chartData);
            return new Cx.ChartSpace(root.ToString(SaveOptions.DisableFormatting));
        }
    }
}
