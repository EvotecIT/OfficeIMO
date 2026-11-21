using static OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTextReader;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static OfficeChartLayout ReadLayout(C.Chart chart, OfficeChartType kind, OfficeChartData data, string? axisTitleFont = null, A.ColorScheme? scheme = null,
            Result? source = null) {
            if (chart.PlotArea?.Descendants<C.Title>().Any(title =>
                    title.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null ||
                    title.GetFirstChild<C.Overlay>() is C.Overlay overlay && overlay.Val?.Value != false) == true)
                OfficeOpenXmlChartFidelityScope.Approximate("Manual axis-title layouts and overlaid axis titles cannot be projected.",
                    "Manually placed or overlaid axis titles are drawn in their automatic positions.");
            var labels = ReadLabels(chart, scheme);
            labels.ResolveSelection(source?.Series.Select(series => series.SourceIndex).ToArray() ??
                    Enumerable.Range(0, data.Series.Count).Select(index => (uint)index).ToArray(),
                out int[]? labeledSeries, out var labeledPoints, out var hiddenLabelPoints);
            var defaultText = ReadTextOrApproximate(() => ReadNativeText(chart, chart, scheme), "chart");
            var legendText = ReadTextOrApproximate(() => ReadNativeText(chart, chart.GetFirstChild<C.Legend>(), scheme), "legend");
            var axisText = ReadTextOrApproximate(() => ReadUniformNativeText(chart, TextAxes(chart), scheme), "axis label");
            var axisTitleText = ReadTextOrApproximate(() => ReadUniformNativeText(chart,
                TextAxes(chart).Select(axis => axis.GetFirstChild<C.Title>()).Where(title => title != null).Cast<OpenXmlElement>(), scheme), "axis title");
            var legend = chart.GetFirstChild<C.Legend>();
            var position = legend?.GetFirstChild<C.LegendPosition>()?.Val?.Value;
            var sharedPosition = position == C.LegendPositionValues.Left ? OfficeChartLegendPosition.Left :
                position == C.LegendPositionValues.Bottom ? OfficeChartLegendPosition.Bottom :
                position == C.LegendPositionValues.Top ? OfficeChartLegendPosition.Top :
                position == C.LegendPositionValues.TopRight ? OfficeChartLegendPosition.TopRight : OfficeChartLegendPosition.Right;
            // Pie and doughnut legends, and those of a single series coloured by point, list categories.
            bool radial = kind == OfficeChartType.Pie || kind == OfficeChartType.Doughnut;
            bool categoryLegend = radial || OfficeOpenXmlChartPlotSemantics.GetPointColoredSeries(chart.PlotArea) != null;
            var hidden = categoryLegend ? legend?.Elements<C.LegendEntry>().Where(entry => entry.GetFirstChild<C.Delete>() is C.Delete delete && delete.Val?.Value != false)
                .Select(entry => entry.GetFirstChild<C.Index>()?.Val?.Value).Where(value => value.HasValue && value.Value <= int.MaxValue)
                .Select(value => (int)value!.Value).ToArray() : null;
            var plot = chart.PlotArea;
            if (plot?.Descendants().Any(element => element is C.TickLabelSkip or C.TickMarkSkip) == true)
                OfficeOpenXmlChartFidelityScope.Approximate("Category-axis label and tick skipping cannot be projected.",
                    "Category-axis labels and tick marks are drawn without skipping.");
            if (plot?.Descendants<C.LabelAlignment>().Any(alignment => alignment.Val?.Value is C.LabelAlignmentValues value && value != C.LabelAlignmentValues.Center) == true)
                OfficeOpenXmlChartFidelityScope.Approximate("Non-centered category label alignment cannot be projected.",
                    "Category labels are drawn centred.");
            var chartLayers = plot?.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToArray();
            // One shared layout places every category point, so all category value axes must agree.
            if (chartLayers?.Any(layer => layer is not C.ScatterChart and not C.BubbleChart) == true && plot!.Elements<C.ValueAxis>()
                .Select(axis => axis.GetFirstChild<C.CrossBetween>()?.Val?.Value ?? C.CrossBetweenValues.Between)
                .Distinct().Skip(1).Any())
                throw new NotSupportedException("Independent cross-between geometry on different value axes cannot be projected.");
            if (plot?.Descendants<C.LabelOffset>().Any(offset => offset.Val?.Value is ushort value && value != 100) == true)
                OfficeOpenXmlChartFidelityScope.Approximate("Independent category label offsets cannot be projected.",
                    "Category labels are drawn at the default distance from their axis.");
            OpenXmlCompositeElement? horizontal = null, vertical = null;
            if (!radial && plot != null) {
                var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
                var layer = plot.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(element =>
                    element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Primary);
                if (layer != null) {
                    var axes = layer.Elements<C.AxisId>().Select(reference => groups.Resolve(reference.Val?.Value)).ToArray();
                    bool numeric = kind == OfficeChartType.Scatter || kind == OfficeChartType.Bubble;
                    if (axes.Length != 2 || axes.Any(axis => axis == null) || ReferenceEquals(axes[0], axes[1]))
                        throw new NotSupportedException("Chart axes must resolve to two distinct supported axes.");
                    horizontal = numeric ? axes[0] : axes.SingleOrDefault(axis => axis is C.CategoryAxis || axis is C.DateAxis);
                    vertical = numeric ? axes[1] : axes.SingleOrDefault(axis => axis is C.ValueAxis);
                    if (horizontal == null || vertical == null || (numeric && (horizontal is not C.ValueAxis || vertical is not C.ValueAxis)))
                        throw new NotSupportedException("The chart axis roles cannot be projected.");
                }
            }
            if (kind == OfficeChartType.Radar &&
                (horizontal?.GetFirstChild<C.Title>() != null || vertical?.GetFirstChild<C.Title>() != null ||
                 vertical?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.MinAxisValue>() != null ||
                 vertical?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.MaxAxisValue>() != null))
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
            return new OfficeChartLayout(overlayLegend: legend?.GetFirstChild<C.Overlay>() is C.Overlay overlay && overlay.Val?.Value != false,
                legendFontSize: legendText.Size, legendFontFamily: legendText.Family,
                legendFontStyle: legendText.Style,
                axisLabelFontSize: axisText.Size, axisTextFontFamily: axisText.Family,
                axisTextFontStyle: axisText.Style,
                axisTitleFontSize: axisTitleText.Size, axisTitleFontStyle: axisTitleText.Style,
                dataLabelFontSize: labels.FontSize ?? defaultText.Size,
                dataLabelFontFamily: labels.FontFamily ?? defaultText.Family,
                dataLabelFontStyle: labels.FontStyle ?? defaultText.Style,
                overlayTitle: chart.GetFirstChild<C.Title>()?.GetFirstChild<C.Overlay>() is C.Overlay title && title.Val?.Value != false,
                showLegend: legend != null, legendPosition: sharedPosition, hiddenCategoryLegendIndexes: hidden,
                showDataLabels: labels.Visible, showDataLabelValues: labels.Values, showDataLabelCategoryNames: labels.Categories,
                showDataLabelSeriesNames: labels.SeriesNames, showDataLabelPercentages: labels.Percentages,
                dataLabelSeparator: labels.Separator, dataLabelNumberFormat: labels.NumberFormat, dataLabelPosition: labels.Position,
                dataLabelSeriesIndexes: labeledSeries, dataLabelPointIndexes: labeledPoints, hiddenDataLabelPointIndexes: hiddenLabelPoints,
                fillRadarSeries: chart.PlotArea?.GetFirstChild<C.RadarChart>()?.RadarStyle?.Val?.Value == C.RadarStyleValues.Filled,
                categoryAxisTitle: ReadLayoutTitle(categoryAxis), valueAxisTitle: ReadLayoutTitle(valueAxis), axisTitleFontFamily: axisTitleFont,
                categoryAxisNumberFormat: categoryAxis is C.ValueAxis ? null : ReadLayoutFormat(categoryAxis, IsClearlyTextualCategoryCache(plot)),
                horizontalAxisNumberFormat: horizontal is C.ValueAxis ? ReadLayoutFormat(horizontal) : null,
                verticalAxisNumberFormat: ReadLayoutFormat(vertical),
                horizontalAxisMinimum: ReadLayoutMinimum(horizontal), horizontalAxisMaximum: ReadLayoutMaximum(horizontal),
                verticalAxisMinimum: ReadLayoutMinimum(vertical), verticalAxisMaximum: ReadLayoutMaximum(vertical),
                horizontalAxisMajorUnit: horizontal?.GetFirstChild<C.MajorUnit>()?.Val?.Value,
                horizontalAxisMinorUnit: horizontal?.GetFirstChild<C.MinorUnit>()?.Val?.Value,
                verticalAxisMajorUnit: vertical?.GetFirstChild<C.MajorUnit>()?.Val?.Value,
                verticalAxisMinorUnit: vertical?.GetFirstChild<C.MinorUnit>()?.Val?.Value,
                showCategoryAxis: !IsDeletedAxis(categoryAxis), showValueAxis: !IsDeletedAxis(valueAxis),
                showCategoryAxisLine: !IsDeletedAxis(categoryAxis), showValueAxisLine: !IsDeletedAxis(valueAxis),
                showCategoryAxisLabels: categoryAxis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value != C.TickLabelPositionValues.None,
                showValueAxisLabels: valueAxis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value != C.TickLabelPositionValues.None,
                horizontalAxisTickLabelPosition: ReadLayoutTickLabelPosition(horizontal),
                verticalAxisTickLabelPosition: ReadLayoutTickLabelPosition(vertical),
                horizontalAxisMajorTickMark: ReadLayoutTick(horizontal?.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
                horizontalAxisMinorTickMark: ReadLayoutTick(horizontal?.GetFirstChild<C.MinorTickMark>()?.Val?.Value),
                verticalAxisMajorTickMark: ReadLayoutTick(vertical?.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
                verticalAxisMinorTickMark: ReadLayoutTick(vertical?.GetFirstChild<C.MinorTickMark>()?.Val?.Value),
                reverseCategoryAxis: categoryAxis is not C.ValueAxis && categoryAxis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin,
                categoryAxisOrientationSpecified: categoryAxis is not C.ValueAxis && categoryAxis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>() != null)
                .WithSecondaryValueAxis(OfficeOpenXmlChartSecondaryAxis.Read(plot))
                .WithDataLabelLeaderLines(labels.LeaderLines)
                .WithCategoryGeometry(plot, scheme)
                .WithPlotSemantics(chart);
        }

        private static void QualifySecondaryLayout(C.PlotArea? plot) {
            if (plot == null) return;
            var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
            var secondaryLayers = plot.ChildElements.OfType<OpenXmlCompositeElement>().Where(element =>
                element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Secondary).ToArray();
            if (secondaryLayers.Length == 0) return;
            // Numeric scales and formats are independent; appearance still uses shared roles.
            foreach (var axis in secondaryLayers.SelectMany(layer => layer.Elements<C.AxisId>())
                .Select(reference => groups.Resolve(reference.Val?.Value)).Distinct()) {
                if (axis == null) throw new NotSupportedException("The secondary chart axes cannot be resolved.");
                if (axis is C.CategoryAxis or C.DateAxis && !IsDeletedAxis(axis))
                    throw new NotSupportedException("Visible secondary category axes cannot be projected.");
                // A deleted axis draws nothing, whatever its stored line.
                if (!IsDeletedAxis(axis) && !HasUnpaintedSecondaryAxisShape(axis.GetFirstChild<C.ChartShapeProperties>()) ||
                    axis.GetFirstChild<C.MajorGridlines>() != null || axis.GetFirstChild<C.MinorGridlines>() != null)
                    OfficeOpenXmlChartFidelityScope.Approximate("The secondary axis appearance cannot be projected independently.",
                        "The secondary axis line and gridlines are drawn like the primary axis.");
                OfficeOpenXmlChartSecondaryAxis.QualifyTitle(axis.GetFirstChild<C.Title>());
                if (axis.GetFirstChild<C.Title>()?.Descendants().Any(child =>
                        child is A.EastAsianFont or A.ComplexScriptFont) == true)
                    throw new NotSupportedException("Script-specific secondary axis title fonts cannot be projected.");
                if (axis is C.ValueAxis && (IsDeletedAxis(axis) || axis.GetFirstChild<C.TickLabelPosition>()?.Val?.Value is C.TickLabelPositionValues tickPosition && tickPosition != C.TickLabelPositionValues.NextTo))
                    throw new NotSupportedException("Secondary value-axis visibility cannot be projected independently.");
                if (axis is C.ValueAxis &&
                    (axis.GetFirstChild<C.CrossesAt>() != null || axis.GetFirstChild<C.DisplayUnits>() != null))
                    throw new NotSupportedException("The secondary numeric axis presentation cannot be projected.");
            }
        }

        private static bool HasUnpaintedSecondaryAxisShape(C.ChartShapeProperties? properties) {
            if (properties == null) return true;
            // Office writes a hidden secondary axis line as no fill, no line and an empty effect list.
            A.Outline? outline = properties.GetFirstChild<A.Outline>();
            return !properties.HasAttributes && outline != null &&
                properties.ChildElements.All(child => child is A.Outline || child is A.NoFill || child is A.EffectList && !child.HasChildren) &&
                outline.ChildElements.Count == 1 && outline.GetFirstChild<A.NoFill>() != null;
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

        private static string? ReadLayoutTitle(OpenXmlCompositeElement? axis) {
            var text = axis?.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>();
            return text?.GetFirstChild<C.RichText>()?.InnerText ?? text?.GetFirstChild<C.StringReference>()?.StringCache?.InnerText;
        }
        private static string? ReadLayoutFormat(OpenXmlCompositeElement? axis, bool textualCategories = false) {
            if (axis == null) return null;
            // A string cache with nonnumeric labels has already materialized its display text;
            // a number format on its category axis cannot change those labels.
            if (textualCategories && axis is C.CategoryAxis && HasUnsupportedSharedAxisNumberFormat(axis)) return null;
            return OfficeOpenXmlChartSecondaryAxis.ReadDrawableNumberFormat(axis);
        }

        /// <summary>Reads the display format of numeric categories, resolving a source-linked format from the category cache.</summary>
        internal static string? ReadCategoryAxisNumberFormat(OpenXmlCompositeElement? axis) {
            if (axis is not C.CategoryAxis and not C.DateAxis) return null;
            C.NumberingFormat? numbering = axis.GetFirstChild<C.NumberingFormat>();
            string? format = numbering?.FormatCode?.Value;
            if (numbering?.SourceLinked?.Value == true || string.IsNullOrWhiteSpace(format)) {
                format = (axis.Parent as C.PlotArea)?.Descendants<C.CategoryAxisData>()
                    .Select(categories => categories.GetFirstChild<C.NumberReference>()?.NumberingCache?.FormatCode?.Text)
                    .FirstOrDefault(code => !string.IsNullOrWhiteSpace(code)) ?? format;
            }
            return string.IsNullOrWhiteSpace(format) || string.Equals(format, "General", StringComparison.OrdinalIgnoreCase)
                ? null : format;
        }

        private static bool IsClearlyTextualCategoryCache(C.PlotArea? plot) {
            C.CategoryAxisData[] caches = plot?.Descendants<C.CategoryAxisData>().ToArray() ?? Array.Empty<C.CategoryAxisData>();
            return caches.Length > 0 && caches.All(category => {
                C.StringPoint[] points = category.GetFirstChild<C.StringReference>()?.StringCache?.Elements<C.StringPoint>().ToArray()
                    ?? Array.Empty<C.StringPoint>();
                return points.Length > 0 && points.All(point =>
                    !double.TryParse(point.NumericValue?.Text, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out _));
            });
        }
        private static double? ReadLayoutMinimum(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.MinAxisValue>()?.Val?.Value;
        private static double? ReadLayoutMaximum(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value;
        private static bool IsDeletedAxis(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Delete>() is C.Delete deleted && deleted.Val?.Value != false;
        private static OfficeChartAxisTickLabelPosition ReadLayoutTickLabelPosition(OpenXmlCompositeElement? axis) =>
            axis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value == C.TickLabelPositionValues.High ? OfficeChartAxisTickLabelPosition.High :
            axis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value == C.TickLabelPositionValues.Low ? OfficeChartAxisTickLabelPosition.Low : OfficeChartAxisTickLabelPosition.NextTo;
        private static OfficeChartAxisTickMark ReadLayoutTick(C.TickMarkValues? value) =>
            value == C.TickMarkValues.Inside ? OfficeChartAxisTickMark.Inside : value == C.TickMarkValues.Outside ? OfficeChartAxisTickMark.Outside :
                value == C.TickMarkValues.Cross ? OfficeChartAxisTickMark.Cross : OfficeChartAxisTickMark.None;
        private static System.Collections.Generic.IEnumerable<OpenXmlElement> TextAxes(C.Chart chart) =>
            chart.PlotArea?.ChildElements.Where(axis => axis is C.CategoryAxis or C.ValueAxis or C.DateAxis).Cast<OpenXmlElement>() ?? Enumerable.Empty<OpenXmlElement>();
    }
}
