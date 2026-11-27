using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string ChartExGraphicDataUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";

        /// <summary>Enumerates native and imported ChartEx drawings on this worksheet.</summary>
        public IEnumerable<ExcelModernChart> ModernCharts {
            get => Locking.ExecuteRead(_excelDocument.EnsureLock(), GetModernChartsCore);
        }

        /// <summary>Returns a native or imported ChartEx drawing by non-visual name.</summary>
        public ExcelModernChart? GetModernChart(string name) {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () =>
                GetModernChartsCore().FirstOrDefault(chart =>
                    string.Equals(chart.Name, name, StringComparison.OrdinalIgnoreCase)));
        }

        private ExcelModernChart[] GetModernChartsCore() {
            ExcelPartRef drawingsPart = ExcelPartRef.WorksheetForMetadata(this).Drawings;
            OfficeXmlNode worksheetDrawing = drawingsPart.IsNone ? default : drawingsPart.Root;
            if (worksheetDrawing.IsNone) return Array.Empty<ExcelModernChart>();
            var charts = new List<ExcelModernChart>();
            foreach (OfficeXmlNode frame in worksheetDrawing.Descendants(Xdr, "graphicFrame")) {
                if (!string.Equals(
                        frame.Child(A, "graphic").Child(A, "graphicData").Attribute("uri"),
                        ChartExGraphicDataUri,
                        StringComparison.Ordinal)) continue;
                ExcelPartRef part = TryGetModernChartPart(frame, drawingsPart);
                charts.Add(new ExcelModernChart(
                    frame,
                    drawingsPart,
                    this,
                    part.IsNone ? null : TryExtractModernChartDataRange(part)));
            }
            return charts.ToArray();
        }

        /// <summary>Adds a native ChartEx chart backed by OfficeIMO's shared hidden chart-data owner.</summary>
        public ExcelModernChart AddModernChart(
            ExcelChartData data,
            int row,
            int column,
            ExcelModernChartType chartType,
            string? title = null,
            OfficeLength? width = null,
            OfficeLength? height = null) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            ArgumentNullException.ThrowIfNull(data);
            ValidateModernChartPlacement(row, column, width.Value, height.Value);
            string layout = GetModernChartLayout(chartType);
            ValidateModernChartData(data);
            ValidateModernChartText(title, nameof(title));

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateDuringDirtyMarks();
            ExcelChartDataRange? range = null;
            OfficeXmlNode frame = default;
            ExcelPartRef drawingsPart = default;
            WriteLock(() => {
                ExcelPartRef worksheetPart = ExcelPartRef.WorksheetForMetadata(this);
                OfficeXmlNode drawing = ExcelPartRef.WorksheetDrawingElement(this);
                if (drawing.Exists) {
                    string drawingRelationshipId = drawing.RelationshipAttribute("id") ?? string.Empty;
                    ExcelPartRef existingDrawingsPart = drawingRelationshipId.Length == 0 ? default : worksheetPart.GetPartById(drawingRelationshipId);
                    if (existingDrawingsPart.Kind != ExcelPartKind.Drawings) {
                        throw new InvalidOperationException("Worksheet drawing relationship is missing or invalid.");
                    }
                    drawingsPart = existingDrawingsPart;
                }

                ExcelSheet dataSheet = _excelDocument.GetOrCreateChartDataSheet();
                int startRow = _excelDocument.ReserveChartDataStartRow(
                    dataSheet,
                    GetChartDataPointCount(data) + 1);
                range = dataSheet.WriteChartData(data, startRow, 1);
                if (drawing.IsNone) {
                    drawingsPart = worksheetPart.AddNewPart(ExcelPartKind.Drawings);
                    drawingsPart.SetRootMarkup(WorksheetDrawingRootMarkup);
                    ExcelPartRef.LinkWorksheetDrawing(this, worksheetPart.GetIdOfPart(drawingsPart)!);
                } else if (drawingsPart.Root.IsNone) {
                    drawingsPart.SetRootMarkup(WorksheetDrawingRootMarkup);
                }

                ExcelPartRef chartPart = drawingsPart.AddNewPart(ExcelPartKind.ExtendedChart);
                chartPart.SetRootMarkup(System.Text.Encoding.UTF8.GetBytes(BuildModernChartSpaceMarkup(data, range!, layout, title)));
                chartPart.SaveRoot();
                string relationshipId = drawingsPart.GetIdOfPart(chartPart)!;
                long widthEmu = width.Value.ToEmus();
                long heightEmu = height.Value.ToEmus();
                uint id = NextDrawingId(drawingsPart);
                string name = "Modern Chart " + id.ToString(CultureInfo.InvariantCulture);
                frame = AppendGraphicFrameAnchor(drawingsPart.Root, row, column, widthEmu, heightEmu, id, name,
                    ChartExGraphicDataUri, context => NewChartExReference(context, relationshipId));
                drawingsPart.SaveRoot();
                MarkRequiresSavePreparation();
            });
            return new ExcelModernChart(frame, drawingsPart, this, range!);
        }

        internal static void ValidateModernChartPlacement(int row, int column, OfficeLength width, OfficeLength height) {
            if (row < 1 || row > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1 || column > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
            if (width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
        }

        internal static void ValidateModernChartData(ExcelChartData data) {
            if (data.Series.Count == 0 || data.Categories.Count == 0) {
                throw new ArgumentException("Modern charts require at least one category and one series.", nameof(data));
            }
            if (data.Series.Count >= ExcelA1.MaxColumns) {
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

        private static ExcelPartRef TryGetModernChartPart(OfficeXmlNode frame, ExcelPartRef drawingsPart) {
            OfficeXmlNode reference = frame.Child(A, "graphic").Child(A, "graphicData").Child(ChartExGraphicDataUri, "chart");
            string relationshipId = reference.RelationshipAttribute("id") ?? string.Empty;
            if (relationshipId.Length == 0) return default;
            foreach ((string Id, ExcelPartRef Part) pair in drawingsPart.Parts) {
                if (string.Equals(pair.Id, relationshipId, StringComparison.Ordinal)) {
                    return pair.Part.Kind == ExcelPartKind.ExtendedChart ? pair.Part : default;
                }
            }
            return default;
        }

        private ExcelChartDataRange? TryExtractModernChartDataRange(ExcelPartRef part) {
            OfficeXmlNode root = part.Root;
            List<OfficeXmlNode> data = ExcelDrawingXml.DescendantsNamed(root, "data");
            if (data.Count == 0) return null;

            string? categoryFormula = FirstFormulaOfDimension(data[0], "strDim", "cat");
            if (!ExcelChartUtils.TryParseSheetQualifiedRange(categoryFormula, out string sheetName, out string categoryRange)
                || !_excelDocument.IsOwnedChartDataSheet(sheetName)
                || !ExcelA1.TryParseRange(categoryRange, out int firstRow, out int categoryColumn, out int lastRow, out int lastCategoryColumn)
                || categoryColumn != lastCategoryColumn) return null;

            for (int index = 0; index < data.Count; index++) {
                string? valueFormula = FirstFormulaOfDimension(data[index], "numDim", "val");
                if (!ExcelChartUtils.TryParseSheetQualifiedRange(valueFormula, out string valueSheet, out string valueRange)
                    || !string.Equals(sheetName, valueSheet, StringComparison.OrdinalIgnoreCase)
                    || !ExcelA1.TryParseRange(valueRange, out int valueFirstRow, out int valueColumn, out int valueLastRow, out int valueLastColumn)
                    || valueFirstRow != firstRow
                    || valueLastRow != lastRow
                    || valueColumn != categoryColumn + index + 1
                    || valueLastColumn != valueColumn) return null;
            }

            bool hasHeaderRow = firstRow > 1;
            if (hasHeaderRow) {
                ExcelSheet dataSheet = _excelDocument.GetSheetForLockedOperation(sheetName);
                List<OfficeXmlNode> series = ExcelDrawingXml.Descendants(root, ChartExGraphicDataUri, "series");
                hasHeaderRow = series.Count == data.Count;
                for (int index = 0; hasHeaderRow && index < series.Count; index++) {
                    OfficeXmlNode expected = ExcelDrawingXml.FirstDescendant(series[index], ChartExGraphicDataUri, "v");
                    string expectedName = expected.IsNone ? string.Empty : expected.InnerText;
                    hasHeaderRow = dataSheet.TryGetCellText(firstRow - 1, categoryColumn + index + 1, out string actualName)
                        && string.Equals(expectedName, actualName, StringComparison.Ordinal);
                }
            }

            return new ExcelChartDataRange(
                sheetName,
                hasHeaderRow ? firstRow - 1 : firstRow,
                categoryColumn,
                lastRow - firstRow + 1,
                data.Count,
                hasHeaderRow);
        }

        // The text of the first f under the first descendant dimension of this name and type, in any namespace.
        private static string? FirstFormulaOfDimension(OfficeXmlNode data, string dimensionName, string type) {
            foreach (OfficeXmlNode dimension in ExcelDrawingXml.DescendantsNamed(data, dimensionName)) {
                if (dimension.Attribute("type") != type) continue;
                OfficeXmlNode formula = ExcelDrawingXml.FirstDescendantNamed(dimension, "f");
                return formula.IsNone ? null : formula.InnerText;
            }
            return null;
        }

        internal static string BuildModernChartSpaceMarkup(
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
                        InvariantNumberText.GetExact(series.Values[index])));
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
            return root.ToString(SaveOptions.DisableFormatting);
        }
    }
}
