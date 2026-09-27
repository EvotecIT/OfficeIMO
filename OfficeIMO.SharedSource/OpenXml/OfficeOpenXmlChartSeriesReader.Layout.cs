using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static OfficeChartLayout ReadLayout(C.Chart chart, OfficeChartKind kind, string? axisTitleFont = null, A.ColorScheme? scheme = null) {
            if (chart.PlotArea?.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null ||
                chart.GetFirstChild<C.Title>()?.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null ||
                chart.GetFirstChild<C.Legend>()?.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null ||
                chart.GetFirstChild<C.Legend>()?.GetFirstChild<C.LegendPosition>()?.Val?.Value == C.LegendPositionValues.TopRight ||
                chart.PlotArea?.Descendants<C.Title>().Any(title => title.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null) == true)
                throw new NotSupportedException("Manual chart layouts and top-right legends cannot be projected.");
            var labels = ReadLabels(chart);
            var defaultText = ReadNativeText(chart, chart, scheme);
            var legendText = ReadNativeText(chart, chart.GetFirstChild<C.Legend>(), scheme);
            var axisText = ReadUniformNativeText(chart, TextAxes(chart), scheme);
            var axisTitleText = ReadUniformNativeText(chart, TextAxes(chart).Select(axis => axis.GetFirstChild<C.Title>()).Where(title => title != null).Cast<OpenXmlElement>(), scheme);
            var legend = chart.GetFirstChild<C.Legend>();
            var position = legend?.GetFirstChild<C.LegendPosition>()?.Val?.Value;
            var sharedPosition = position == C.LegendPositionValues.Left ? OfficeChartLegendPosition.Left :
                position == C.LegendPositionValues.Bottom ? OfficeChartLegendPosition.Bottom :
                position == C.LegendPositionValues.Top ? OfficeChartLegendPosition.Top : OfficeChartLegendPosition.Right;
            bool radial = kind == OfficeChartKind.Pie || kind == OfficeChartKind.Doughnut;
            var hidden = radial ? legend?.Elements<C.LegendEntry>().Where(entry => entry.GetFirstChild<C.Delete>() is C.Delete delete && delete.Val?.Value != false)
                .Select(entry => entry.GetFirstChild<C.Index>()?.Val?.Value).Where(value => value.HasValue && value.Value <= int.MaxValue)
                .Select(value => (int)value!.Value).ToArray() : null;
            var plot = chart.PlotArea;
            if (plot?.Descendants().Any(element => element is C.TickLabelSkip or C.TickMarkSkip) == true)
                throw new NotSupportedException("Category-axis label and tick skipping cannot be projected.");
            if (plot?.Descendants<C.LabelAlignment>().Any(alignment => alignment.Val?.Value is C.LabelAlignmentValues value && value != C.LabelAlignmentValues.Center) == true)
                throw new NotSupportedException("Non-centered category label alignment cannot be projected.");
            if (plot?.Descendants<C.LabelOffset>().Any(offset => offset.Val?.Value is ushort value && value != 100) == true ||
                plot?.Descendants<C.CrossBetween>().Any(crossing => crossing.Val?.Value is C.CrossBetweenValues value && value != C.CrossBetweenValues.Between) == true)
                throw new NotSupportedException("Independent category label offsets and cross-between geometry cannot be projected.");
            OpenXmlCompositeElement? horizontal = null, vertical = null;
            if (!radial && plot != null) {
                var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
                var layer = plot.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(element =>
                    element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Primary);
                if (layer != null) {
                    var axes = layer.Elements<C.AxisId>().Select(reference => groups.Resolve(reference.Val?.Value)).ToArray();
                    bool numeric = kind == OfficeChartKind.Scatter || kind == OfficeChartKind.Bubble;
                    if (axes.Length != 2 || axes.Any(axis => axis == null) || ReferenceEquals(axes[0], axes[1]))
                        throw new NotSupportedException("Chart axes must resolve to two distinct supported axes.");
                    horizontal = numeric ? axes[0] : axes.SingleOrDefault(axis => axis is C.CategoryAxis || axis is C.DateAxis);
                    vertical = numeric ? axes[1] : axes.SingleOrDefault(axis => axis is C.ValueAxis);
                    if (horizontal == null || vertical == null || (numeric && (horizontal is not C.ValueAxis || vertical is not C.ValueAxis)))
                        throw new NotSupportedException("The chart axis roles cannot be projected.");
                }
            }
            foreach (var axis in new[] { horizontal, vertical }.Where(axis => axis != null)) {
                if (axis!.GetFirstChild<C.Scaling>()?.GetFirstChild<C.LogBase>() != null)
                    throw new NotSupportedException("Logarithmic chart axes cannot be projected.");
                if (axis is C.ValueAxis && axis.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin)
                    throw new NotSupportedException("Reversed numeric chart axes cannot be projected.");
                if (axis is C.ValueAxis numericAxis && HasUnsupportedSharedAxisNumberFormat(numericAxis))
                    throw new NotSupportedException("The native numeric axis format cannot be projected.");
                if (axis.GetFirstChild<C.CrossesAt>() != null || axis.GetFirstChild<C.DisplayUnits>() != null)
                    throw new NotSupportedException("Explicit axis crossing values and display units require an independent axis projection.");
                if (kind is OfficeChartKind.BarClustered or OfficeChartKind.BarStacked or OfficeChartKind.BarStacked100 &&
                    axis.GetFirstChild<C.Crosses>()?.Val?.Value == C.CrossesValues.Maximum)
                    throw new NotSupportedException("Maximum crossing for horizontal bar axes cannot be projected.");
            }
            QualifySecondaryLayout(plot, vertical);
            // Titles, visibility and category direction describe logical roles;
            // scales and tick marks describe the physical horizontal/vertical axes.
            var categoryAxis = horizontal;
            var valueAxis = vertical;
            if (kind == OfficeChartKind.BarClustered || kind == OfficeChartKind.BarStacked || kind == OfficeChartKind.BarStacked100) {
                horizontal = valueAxis;
                vertical = categoryAxis;
            }
            return new OfficeChartLayout(overlayLegend: legend?.GetFirstChild<C.Overlay>() is C.Overlay overlay && overlay.Val?.Value != false,
                legendFontSize: legendText.Size, legendFontStyle: legendText.Style,
                axisLabelFontSize: axisText.Size, axisTextFontStyle: axisText.Style,
                axisTitleFontSize: axisTitleText.Size, axisTitleFontStyle: axisTitleText.Style,
                dataLabelFontSize: defaultText.Size, dataLabelFontStyle: defaultText.Style,
                overlayTitle: chart.GetFirstChild<C.Title>()?.GetFirstChild<C.Overlay>() is C.Overlay title && title.Val?.Value != false,
                showLegend: legend != null, legendPosition: sharedPosition, hiddenCategoryLegendIndexes: hidden,
                showDataLabels: labels.Visible, showDataLabelValues: labels.Values, showDataLabelCategoryNames: labels.Categories,
                showDataLabelSeriesNames: labels.SeriesNames, showDataLabelPercentages: labels.Percentages,
                dataLabelSeparator: labels.Separator, dataLabelNumberFormat: labels.NumberFormat, dataLabelPosition: labels.Position,
                fillRadarSeries: chart.PlotArea?.GetFirstChild<C.RadarChart>()?.RadarStyle?.Val?.Value == C.RadarStyleValues.Filled,
                categoryAxisTitle: ReadLayoutTitle(categoryAxis), valueAxisTitle: ReadLayoutTitle(valueAxis), axisTitleFontFamily: axisTitleFont,
                categoryAxisNumberFormat: categoryAxis is C.ValueAxis ? null : ReadLayoutFormat(categoryAxis),
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
                // A native axis's crosses value positions the perpendicular axis.
                horizontalAxisCrossingPosition: ReadLayoutCrossing(vertical), verticalAxisCrossingPosition: ReadLayoutCrossing(horizontal),
                horizontalAxisMajorTickMark: ReadLayoutTick(horizontal?.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
                horizontalAxisMinorTickMark: ReadLayoutTick(horizontal?.GetFirstChild<C.MinorTickMark>()?.Val?.Value),
                verticalAxisMajorTickMark: ReadLayoutTick(vertical?.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
                verticalAxisMinorTickMark: ReadLayoutTick(vertical?.GetFirstChild<C.MinorTickMark>()?.Val?.Value),
                reverseCategoryAxis: categoryAxis is not C.ValueAxis && categoryAxis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin,
                categoryAxisOrientationSpecified: categoryAxis is not C.ValueAxis && categoryAxis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>() != null);
        }

        private static void QualifySecondaryLayout(C.PlotArea? plot, OpenXmlCompositeElement? primaryValueAxis) {
            if (plot == null) return;
            var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
            var secondaryLayers = plot.ChildElements.OfType<OpenXmlCompositeElement>().Where(element =>
                element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Secondary).ToArray();
            if (secondaryLayers.Length == 0) return;
            // The shared secondary renderer currently uses automatic scales and the primary
            // value-label treatment. Reject native settings it cannot represent independently.
            foreach (var axis in secondaryLayers.SelectMany(layer => layer.Elements<C.AxisId>())
                .Select(reference => groups.Resolve(reference.Val?.Value)).Distinct()) {
                if (axis == null) throw new NotSupportedException("The secondary chart axes cannot be resolved.");
                if (axis is C.CategoryAxis or C.DateAxis && !IsDeletedAxis(axis))
                    throw new NotSupportedException("Visible secondary category axes cannot be projected.");
                if (axis.GetFirstChild<C.Title>() != null || axis.GetFirstChild<C.ChartShapeProperties>() != null ||
                    axis.GetFirstChild<C.MajorGridlines>() != null || axis.GetFirstChild<C.MinorGridlines>() != null)
                    throw new NotSupportedException("The secondary axis appearance cannot be projected independently.");
                if (axis is C.ValueAxis && (IsDeletedAxis(axis) || axis.GetFirstChild<C.TickLabelPosition>()?.Val?.Value is C.TickLabelPositionValues tickPosition && tickPosition != C.TickLabelPositionValues.NextTo))
                    throw new NotSupportedException("Secondary value-axis visibility cannot be projected independently.");
                if (axis is C.ValueAxis) QualifyAutomaticSecondaryScale(axis);
            }
            if (primaryValueAxis != null) QualifyAutomaticSecondaryScale(primaryValueAxis);
        }

        private static void QualifyAutomaticSecondaryScale(OpenXmlCompositeElement axis) {
            var scaling = axis.GetFirstChild<C.Scaling>();
            if (scaling?.GetFirstChild<C.MinAxisValue>() != null || scaling?.GetFirstChild<C.MaxAxisValue>() != null ||
                scaling?.GetFirstChild<C.LogBase>() != null || scaling?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin ||
                axis.GetFirstChild<C.MajorUnit>() != null || axis.GetFirstChild<C.MinorUnit>() != null ||
                axis.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value is string format && !string.Equals(format, "General", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Independent secondary-axis scales and formats cannot be projected.");
        }

        private static string? ReadLayoutTitle(OpenXmlCompositeElement? axis) {
            var text = axis?.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>();
            return text?.GetFirstChild<C.RichText>()?.InnerText ?? text?.GetFirstChild<C.StringReference>()?.StringCache?.InnerText;
        }
        private static string? ReadLayoutFormat(OpenXmlCompositeElement? axis) {
            if (axis != null && HasUnsupportedSharedAxisNumberFormat(axis))
                throw new NotSupportedException("The native axis format cannot be projected.");
            return axis?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
        }
        private static double? ReadLayoutMinimum(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.MinAxisValue>()?.Val?.Value;
        private static double? ReadLayoutMaximum(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value;
        private static bool IsDeletedAxis(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Delete>() is C.Delete deleted && deleted.Val?.Value != false;
        private static OfficeChartAxisTickLabelPosition ReadLayoutTickLabelPosition(OpenXmlCompositeElement? axis) =>
            axis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value == C.TickLabelPositionValues.High ? OfficeChartAxisTickLabelPosition.High :
            axis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value == C.TickLabelPositionValues.Low ? OfficeChartAxisTickLabelPosition.Low : OfficeChartAxisTickLabelPosition.NextTo;
        private static OfficeChartAxisCrossingPosition ReadLayoutCrossing(OpenXmlCompositeElement? axis) =>
            axis?.GetFirstChild<C.Crosses>()?.Val?.Value == C.CrossesValues.Maximum ? OfficeChartAxisCrossingPosition.Maximum :
            axis?.GetFirstChild<C.Crosses>()?.Val?.Value == C.CrossesValues.Minimum ? OfficeChartAxisCrossingPosition.Minimum : OfficeChartAxisCrossingPosition.AutoZero;
        private static OfficeChartAxisTickMark ReadLayoutTick(C.TickMarkValues? value) =>
            value == C.TickMarkValues.Inside ? OfficeChartAxisTickMark.Inside : value == C.TickMarkValues.Outside ? OfficeChartAxisTickMark.Outside :
                value == C.TickMarkValues.Cross ? OfficeChartAxisTickMark.Cross : OfficeChartAxisTickMark.None;
    }
}
