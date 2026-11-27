using static OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTextReader;
using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static OfficeChartLayout ReadLayout(OfficeXmlNode chart, OfficeChartType kind, OfficeChartData data, string? axisTitleFont = null, OfficeXmlNode scheme = default,
            Result? source = null) {
            OfficeXmlNode plot = chart.Child(NsC, "plotArea");
            if (plot.Descendants(NsC, "title").Any(title =>
                    !title.Child(NsC, "layout").Child(NsC, "manualLayout").IsNone ||
                    title.Child(NsC, "overlay") is { IsNone: false } overlay && overlay.BoolVal() != false))
                OfficeOpenXmlChartFidelityScope.Approximate("Manual axis-title layouts and overlaid axis titles cannot be projected.",
                    "Manually placed or overlaid axis titles are drawn in their automatic positions.");
            var labels = ReadLabels(chart, scheme);
            labels.ResolveSelection(source?.Series.Select(series => series.SourceIndex).ToArray() ??
                    Enumerable.Range(0, data.Series.Count).Select(index => (uint)index).ToArray(),
                out int[]? labeledSeries, out var labeledPoints, out var hiddenLabelPoints);
            var defaultText = ReadTextOrApproximate(() => ReadNativeText(chart, chart, scheme), "chart");
            var legendText = ReadTextOrApproximate(() => ReadNativeText(chart, chart.Child(NsC, "legend"), scheme), "legend");
            var axisText = ReadTextOrApproximate(() => ReadUniformNativeText(chart, TextAxes(chart), scheme), "axis label");
            var axisTitleText = ReadTextOrApproximate(() => ReadUniformNativeText(chart,
                TextAxes(chart).Select(axis => axis.Child(NsC, "title")).Where(title => !title.IsNone), scheme), "axis title");
            var legend = chart.Child(NsC, "legend");
            string? position = legend.ChildVal(NsC, "legendPos");
            var sharedPosition = position == "l" ? OfficeChartLegendPosition.Left :
                position == "b" ? OfficeChartLegendPosition.Bottom :
                position == "t" ? OfficeChartLegendPosition.Top :
                position == "tr" ? OfficeChartLegendPosition.TopRight : OfficeChartLegendPosition.Right;
            // Pie and doughnut legends, and those of a single series coloured by point, list categories.
            bool radial = kind == OfficeChartType.Pie || kind == OfficeChartType.Doughnut;
            bool categoryLegend = radial || !OfficeOpenXmlChartPlotSemantics.GetPointColoredSeries(plot).IsNone;
            var hidden = categoryLegend && !legend.IsNone ? legend.Children(NsC, "legendEntry").ToList()
                .Where(entry => entry.Child(NsC, "delete") is { IsNone: false } delete && delete.BoolVal() != false)
                .Select(entry => entry.Child(NsC, "idx").UIntVal()).Where(value => value.HasValue && value.Value <= int.MaxValue)
                .Select(value => (int)value!.Value).ToArray() : null;
            if (!plot.FirstDescendant(NsC, "tickLblSkip").IsNone || !plot.FirstDescendant(NsC, "tickMarkSkip").IsNone)
                OfficeOpenXmlChartFidelityScope.Approximate("Category-axis label and tick skipping cannot be projected.",
                    "Category-axis labels and tick marks are drawn without skipping.");
            if (plot.Descendants(NsC, "lblAlgn").Any(alignment => alignment.Val() is string value && value != "ctr"))
                OfficeOpenXmlChartFidelityScope.Approximate("Non-centered category label alignment cannot be projected.",
                    "Category labels are drawn centred.");
            var chartLayers = plot.IsNone ? null : plot.Children().ToList()
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToArray();
            // One shared layout places every category point, so all category value axes must agree.
            if (chartLayers?.Any(layer => !layer.IsChart("scatterChart") && !layer.IsChart("bubbleChart")) == true && plot.Children(NsC, "valAx").ToList()
                .Select(axis => axis.ChildVal(NsC, "crossBetween") ?? "between")
                .Distinct().Skip(1).Any())
                throw new NotSupportedException("Independent cross-between geometry on different value axes cannot be projected.");
            if (plot.Descendants(NsC, "lblOffset").Any(offset => offset.UShortVal() is ushort value && value != 100))
                OfficeOpenXmlChartFidelityScope.Approximate("Independent category label offsets cannot be projected.",
                    "Category labels are drawn at the default distance from their axis.");
            OfficeXmlNode horizontal = OfficeXmlNode.None, vertical = OfficeXmlNode.None;
            if (!radial && !plot.IsNone) {
                var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
                var layer = plot.Children().ToList().FirstOrDefault(element =>
                    element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Primary);
                if (!layer.IsNone) {
                    var axes = layer.Children(NsC, "axId").ToList().Select(reference => groups.Resolve(reference.UIntVal())).ToArray();
                    bool numeric = kind == OfficeChartType.Scatter || kind == OfficeChartType.Bubble;
                    if (axes.Length != 2 || axes.Any(axis => axis.IsNone) || axes[0] == axes[1])
                        throw new NotSupportedException("Chart axes must resolve to two distinct supported axes.");
                    horizontal = numeric ? axes[0] : axes.SingleOrDefault(axis => axis.IsChart("catAx") || axis.IsChart("dateAx"));
                    vertical = numeric ? axes[1] : axes.SingleOrDefault(axis => axis.IsChart("valAx"));
                    if (horizontal.IsNone || vertical.IsNone || (numeric && (!horizontal.IsChart("valAx") || !vertical.IsChart("valAx"))))
                        throw new NotSupportedException("The chart axis roles cannot be projected.");
                }
            }
            if (kind == OfficeChartType.Radar &&
                (!horizontal.Child(NsC, "title").IsNone || !vertical.Child(NsC, "title").IsNone ||
                 !vertical.Child(NsC, "scaling").Child(NsC, "min").IsNone ||
                 !vertical.Child(NsC, "scaling").Child(NsC, "max").IsNone))
                throw new NotSupportedException("Radar axis titles and explicit scales cannot be projected.");
            QualifySecondaryLayout(plot);
            OfficeOpenXmlChartSecondaryAxis.QualifyLinearProjection(plot);
            // Titles, visibility and category direction describe logical roles;
            // scales and tick marks describe the physical horizontal/vertical axes.
            var categoryAxis = horizontal;
            var valueAxis = vertical;
            if (kind == OfficeChartType.BarClustered || kind == OfficeChartType.BarStacked || kind == OfficeChartType.BarStacked100) {
                horizontal = valueAxis;
                vertical = categoryAxis;
            }
            return new OfficeChartLayout(overlayLegend: legend.Child(NsC, "overlay") is { IsNone: false } overlay && overlay.BoolVal() != false,
                legendFontSize: legendText.Size, legendFontFamily: legendText.Family,
                legendFontStyle: legendText.Style,
                axisLabelFontSize: axisText.Size, axisTextFontFamily: axisText.Family,
                axisTextFontStyle: axisText.Style,
                axisTitleFontSize: axisTitleText.Size, axisTitleFontStyle: axisTitleText.Style,
                dataLabelFontSize: labels.FontSize ?? defaultText.Size,
                dataLabelFontFamily: labels.FontFamily ?? defaultText.Family,
                dataLabelFontStyle: labels.FontStyle ?? defaultText.Style,
                overlayTitle: chart.Child(NsC, "title").Child(NsC, "overlay") is { IsNone: false } title && title.BoolVal() != false,
                showLegend: !legend.IsNone, legendPosition: sharedPosition, hiddenCategoryLegendIndexes: hidden,
                showDataLabels: labels.Visible, showDataLabelValues: labels.Values, showDataLabelCategoryNames: labels.Categories,
                showDataLabelSeriesNames: labels.SeriesNames, showDataLabelPercentages: labels.Percentages,
                dataLabelSeparator: labels.Separator, dataLabelNumberFormat: labels.NumberFormat, dataLabelPosition: labels.Position,
                dataLabelSeriesIndexes: labeledSeries, dataLabelPointIndexes: labeledPoints, hiddenDataLabelPointIndexes: hiddenLabelPoints,
                fillRadarSeries: plot.Child(NsC, "radarChart").ChildVal(NsC, "radarStyle") == "filled",
                categoryAxisTitle: ReadLayoutTitle(categoryAxis), valueAxisTitle: ReadLayoutTitle(valueAxis), axisTitleFontFamily: axisTitleFont,
                categoryAxisNumberFormat: categoryAxis.IsChart("valAx") ? null : ReadLayoutFormat(categoryAxis, IsClearlyTextualCategoryCache(plot)),
                horizontalAxisNumberFormat: horizontal.IsChart("valAx") ? ReadLayoutFormat(horizontal) : null,
                verticalAxisNumberFormat: ReadLayoutFormat(vertical),
                horizontalAxisMinimum: ReadLayoutMinimum(horizontal), horizontalAxisMaximum: ReadLayoutMaximum(horizontal),
                verticalAxisMinimum: ReadLayoutMinimum(vertical), verticalAxisMaximum: ReadLayoutMaximum(vertical),
                horizontalAxisMajorUnit: horizontal.Child(NsC, "majorUnit").DoubleVal(),
                horizontalAxisMinorUnit: horizontal.Child(NsC, "minorUnit").DoubleVal(),
                verticalAxisMajorUnit: vertical.Child(NsC, "majorUnit").DoubleVal(),
                verticalAxisMinorUnit: vertical.Child(NsC, "minorUnit").DoubleVal(),
                showCategoryAxis: !IsDeletedAxis(categoryAxis), showValueAxis: !IsDeletedAxis(valueAxis),
                showCategoryAxisLine: !IsDeletedAxis(categoryAxis), showValueAxisLine: !IsDeletedAxis(valueAxis),
                showCategoryAxisLabels: categoryAxis.ChildVal(NsC, "tickLblPos") != "none",
                showValueAxisLabels: valueAxis.ChildVal(NsC, "tickLblPos") != "none",
                horizontalAxisTickLabelPosition: ReadLayoutTickLabelPosition(horizontal),
                verticalAxisTickLabelPosition: ReadLayoutTickLabelPosition(vertical),
                horizontalAxisMajorTickMark: ReadLayoutTick(horizontal.ChildVal(NsC, "majorTickMark")),
                horizontalAxisMinorTickMark: ReadLayoutTick(horizontal.ChildVal(NsC, "minorTickMark")),
                verticalAxisMajorTickMark: ReadLayoutTick(vertical.ChildVal(NsC, "majorTickMark")),
                verticalAxisMinorTickMark: ReadLayoutTick(vertical.ChildVal(NsC, "minorTickMark")),
                reverseCategoryAxis: !categoryAxis.IsChart("valAx") && categoryAxis.Child(NsC, "scaling").ChildVal(NsC, "orientation") == "maxMin",
                categoryAxisOrientationSpecified: !categoryAxis.IsChart("valAx") && !categoryAxis.Child(NsC, "scaling").Child(NsC, "orientation").IsNone)
                .WithSecondaryValueAxis(OfficeOpenXmlChartSecondaryAxis.Read(plot))
                .WithDataLabelLeaderLines(labels.LeaderLines)
                .WithCategoryGeometry(plot, scheme)
                .WithPlotSemantics(chart);
        }

        private static void QualifySecondaryLayout(OfficeXmlNode plot) {
            if (plot.IsNone) return;
            var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
            var secondaryLayers = plot.Children().ToList().Where(element =>
                element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Secondary).ToArray();
            if (secondaryLayers.Length == 0) return;
            // Numeric scales and formats are independent; appearance still uses shared roles.
            foreach (var axis in secondaryLayers.SelectMany(layer => layer.Children(NsC, "axId").ToList())
                .Select(reference => groups.Resolve(reference.UIntVal())).Distinct()) {
                if (axis.IsNone) throw new NotSupportedException("The secondary chart axes cannot be resolved.");
                if ((axis.IsChart("catAx") || axis.IsChart("dateAx")) && !IsDeletedAxis(axis))
                    throw new NotSupportedException("Visible secondary category axes cannot be projected.");
                // A deleted axis draws nothing, whatever its stored line.
                if (!IsDeletedAxis(axis) && !HasUnpaintedSecondaryAxisShape(axis.Child(NsC, "spPr")) ||
                    !axis.Child(NsC, "majorGridlines").IsNone || !axis.Child(NsC, "minorGridlines").IsNone)
                    OfficeOpenXmlChartFidelityScope.Approximate("The secondary axis appearance cannot be projected independently.",
                        "The secondary axis line and gridlines are drawn like the primary axis.");
                OfficeOpenXmlChartSecondaryAxis.QualifyTitle(axis.Child(NsC, "title"));
                OfficeXmlNode axisTitle = axis.Child(NsC, "title");
                if (!axisTitle.IsNone && axisTitle.AllDescendants().Any(child =>
                        child.IsDrawing("ea") || child.IsDrawing("cs")))
                    throw new NotSupportedException("Script-specific secondary axis title fonts cannot be projected.");
                if (axis.IsChart("valAx") && (IsDeletedAxis(axis) || axis.ChildVal(NsC, "tickLblPos") is string tickPosition && tickPosition != "nextTo"))
                    throw new NotSupportedException("Secondary value-axis visibility cannot be projected independently.");
                if (axis.IsChart("valAx") &&
                    (!axis.Child(NsC, "crossesAt").IsNone || !axis.Child(NsC, "dispUnits").IsNone))
                    throw new NotSupportedException("The secondary numeric axis presentation cannot be projected.");
            }
        }

        private static bool HasUnpaintedSecondaryAxisShape(OfficeXmlNode properties) {
            if (properties.IsNone) return true;
            // Office writes a hidden secondary axis line as no fill, no line and an empty effect list.
            OfficeXmlNode outline = properties.Child(NsA, "ln");
            return !properties.HasAttributes && !outline.IsNone &&
                properties.Children().ToList().All(child => child.IsDrawing("ln") || child.IsDrawing("noFill") || child.IsDrawing("effectLst") && !child.HasChildElements()) &&
                outline.Children().Count == 1 && !outline.Child(NsA, "noFill").IsNone;
        }

        /// <summary>
        /// Reads chart text formatting. Strict projections reject text the renderer cannot reproduce; best-effort
        /// projections draw that text with the default formatting and report it.
        /// </summary>
        private static NativeText ReadTextOrApproximate(Func<NativeText> read, string role) {
            int checkpoint = OfficeOpenXmlChartFidelityScope.Checkpoint;
            try {
                return read();
            } catch (NotSupportedException exception) when (OfficeOpenXmlChartFidelityScope.AllowsApproximation) {
                OfficeOpenXmlChartFidelityScope.RollBack(checkpoint);
                OfficeOpenXmlChartFidelityScope.Record("The " + role + " text is drawn with default formatting: " + exception.Message);
                return default;
            }
        }

        private static string? ReadLayoutTitle(OfficeXmlNode axis) {
            var text = axis.Child(NsC, "title").Child(NsC, "tx");
            OfficeXmlNode rich = text.Child(NsC, "rich");
            if (!rich.IsNone) return rich.InnerText;
            OfficeXmlNode cache = text.Child(NsC, "strRef").Child(NsC, "strCache");
            return cache.IsNone ? null : cache.InnerText;
        }
        private static string? ReadLayoutFormat(OfficeXmlNode axis, bool textualCategories = false) {
            if (axis.IsNone) return null;
            // A string cache with nonnumeric labels has already materialized its display text;
            // a number format on its category axis cannot change those labels.
            if (textualCategories && axis.IsChart("catAx") && HasUnsupportedSharedAxisNumberFormat(axis)) return null;
            return OfficeOpenXmlChartSecondaryAxis.ReadDrawableNumberFormat(axis);
        }

        /// <summary>Reads the display format of numeric categories, resolving a source-linked format from the category cache.</summary>
        internal static string? ReadCategoryAxisNumberFormat(OfficeXmlNode axis) {
            if (!axis.IsChart("catAx") && !axis.IsChart("dateAx")) return null;
            OfficeXmlNode numbering = axis.Child(NsC, "numFmt");
            string? format = numbering.Attribute("formatCode");
            if (numbering.BoolAttr("sourceLinked") == true || string.IsNullOrWhiteSpace(format)) {
                OfficeXmlNode parent = axis.Parent;
                if (parent.IsChart("plotArea"))
                    format = parent.Descendants(NsC, "cat")
                        .Select(categories => categories.Child(NsC, "numRef").Child(NsC, "numCache").ChildText(NsC, "formatCode"))
                        .FirstOrDefault(code => !string.IsNullOrWhiteSpace(code)) ?? format;
            }
            return string.IsNullOrWhiteSpace(format) || string.Equals(format, "General", StringComparison.OrdinalIgnoreCase)
                ? null : format;
        }

        private static bool IsClearlyTextualCategoryCache(OfficeXmlNode plot) {
            List<OfficeXmlNode> caches = plot.Descendants(NsC, "cat");
            return caches.Count > 0 && caches.All(category => {
                List<OfficeXmlNode> points = category.Child(NsC, "strRef").Child(NsC, "strCache").Children(NsC, "pt").ToList();
                return points.Count > 0 && points.All(point =>
                    !OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(point.ChildText(NsC, "v"), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double _));
            });
        }
        private static double? ReadLayoutMinimum(OfficeXmlNode axis) => axis.Child(NsC, "scaling").Child(NsC, "min").DoubleVal();
        private static double? ReadLayoutMaximum(OfficeXmlNode axis) => axis.Child(NsC, "scaling").Child(NsC, "max").DoubleVal();
        private static bool IsDeletedAxis(OfficeXmlNode axis) => axis.Child(NsC, "delete") is { IsNone: false } deleted && deleted.BoolVal() != false;
        private static OfficeChartAxisTickLabelPosition ReadLayoutTickLabelPosition(OfficeXmlNode axis) =>
            axis.ChildVal(NsC, "tickLblPos") == "high" ? OfficeChartAxisTickLabelPosition.High :
            axis.ChildVal(NsC, "tickLblPos") == "low" ? OfficeChartAxisTickLabelPosition.Low : OfficeChartAxisTickLabelPosition.NextTo;
        private static OfficeChartAxisTickMark ReadLayoutTick(string? value) =>
            value == "in" ? OfficeChartAxisTickMark.Inside : value == "out" ? OfficeChartAxisTickMark.Outside :
                value == "cross" ? OfficeChartAxisTickMark.Cross : OfficeChartAxisTickMark.None;
        private static IEnumerable<OfficeXmlNode> TextAxes(OfficeXmlNode chart) =>
            chart.Child(NsC, "plotArea").Children().ToList().Where(axis => axis.IsChart("catAx") || axis.IsChart("valAx") || axis.IsChart("dateAx"));

        // SDK callers
        internal static OfficeChartLayout ReadLayout(C.Chart chart, OfficeChartType kind, OfficeChartData data, string? axisTitleFont = null, A.ColorScheme? scheme = null,
            Result? source = null) =>
            ReadLayout(OfficeXmlNode.FromSdk(chart), kind, data, axisTitleFont, OfficeXmlNode.FromSdk(scheme), source);

        internal static string? ReadCategoryAxisNumberFormat(OpenXmlCompositeElement? axis) => ReadCategoryAxisNumberFormat(OfficeXmlNode.FromSdk(axis));
    }
}
