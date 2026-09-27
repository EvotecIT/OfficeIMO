using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static OfficeChartLayout ReadLayout(C.Chart chart, OfficeChartKind kind, string? axisTitleFont = null) {
            var labels = ReadLabels(chart);
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
            }
            QualifySecondaryLayout(plot, vertical);
            return new OfficeChartLayout(overlayLegend: legend?.GetFirstChild<C.Overlay>() is C.Overlay overlay && overlay.Val?.Value != false,
                overlayTitle: chart.GetFirstChild<C.Title>()?.GetFirstChild<C.Overlay>() is C.Overlay title && title.Val?.Value != false,
                showLegend: legend != null, legendPosition: sharedPosition, hiddenCategoryLegendIndexes: hidden,
                showDataLabels: labels.Visible, showDataLabelValues: labels.Values, showDataLabelCategoryNames: labels.Categories,
                showDataLabelSeriesNames: labels.SeriesNames, showDataLabelPercentages: labels.Percentages,
                dataLabelSeparator: labels.Separator, dataLabelNumberFormat: labels.NumberFormat, dataLabelPosition: labels.Position,
                fillRadarSeries: chart.PlotArea?.GetFirstChild<C.RadarChart>()?.RadarStyle?.Val?.Value == C.RadarStyleValues.Filled,
                categoryAxisTitle: ReadLayoutTitle(horizontal), valueAxisTitle: ReadLayoutTitle(vertical), axisTitleFontFamily: axisTitleFont,
                categoryAxisNumberFormat: horizontal is C.ValueAxis ? null : ReadLayoutFormat(horizontal),
                horizontalAxisNumberFormat: horizontal is C.ValueAxis ? ReadLayoutFormat(horizontal) : null,
                verticalAxisNumberFormat: ReadLayoutFormat(vertical),
                horizontalAxisMinimum: ReadLayoutMinimum(horizontal), horizontalAxisMaximum: ReadLayoutMaximum(horizontal),
                verticalAxisMinimum: ReadLayoutMinimum(vertical), verticalAxisMaximum: ReadLayoutMaximum(vertical),
                horizontalAxisMajorUnit: horizontal?.GetFirstChild<C.MajorUnit>()?.Val?.Value,
                horizontalAxisMinorUnit: horizontal?.GetFirstChild<C.MinorUnit>()?.Val?.Value,
                verticalAxisMajorUnit: vertical?.GetFirstChild<C.MajorUnit>()?.Val?.Value,
                verticalAxisMinorUnit: vertical?.GetFirstChild<C.MinorUnit>()?.Val?.Value,
                showCategoryAxis: !IsDeletedAxis(horizontal), showValueAxis: !IsDeletedAxis(vertical),
                showCategoryAxisLabels: horizontal?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value != C.TickLabelPositionValues.None,
                showValueAxisLabels: vertical?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value != C.TickLabelPositionValues.None,
                horizontalAxisMajorTickMark: ReadLayoutTick(horizontal?.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
                horizontalAxisMinorTickMark: ReadLayoutTick(horizontal?.GetFirstChild<C.MinorTickMark>()?.Val?.Value),
                verticalAxisMajorTickMark: ReadLayoutTick(vertical?.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
                verticalAxisMinorTickMark: ReadLayoutTick(vertical?.GetFirstChild<C.MinorTickMark>()?.Val?.Value),
                reverseCategoryAxis: horizontal is not C.ValueAxis && horizontal?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin,
                categoryAxisOrientationSpecified: horizontal is not C.ValueAxis && horizontal?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>() != null);
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
                if (axis.GetFirstChild<C.Title>() != null || axis.GetFirstChild<C.ChartShapeProperties>() != null ||
                    axis.GetFirstChild<C.MajorGridlines>() != null || axis.GetFirstChild<C.MinorGridlines>() != null)
                    throw new NotSupportedException("The secondary axis appearance cannot be projected independently.");
                if (axis is C.ValueAxis && (IsDeletedAxis(axis) || axis.GetFirstChild<C.TickLabelPosition>()?.Val?.Value == C.TickLabelPositionValues.None))
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
        private static string? ReadLayoutFormat(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
        private static double? ReadLayoutMinimum(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.MinAxisValue>()?.Val?.Value;
        private static double? ReadLayoutMaximum(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Scaling>()?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value;
        private static bool IsDeletedAxis(OpenXmlCompositeElement? axis) => axis?.GetFirstChild<C.Delete>() is C.Delete deleted && deleted.Val?.Value != false;
        private static OfficeChartAxisTickMark ReadLayoutTick(C.TickMarkValues? value) =>
            value == C.TickMarkValues.Inside ? OfficeChartAxisTickMark.Inside : value == C.TickMarkValues.Outside ? OfficeChartAxisTickMark.Outside :
                value == C.TickMarkValues.Cross ? OfficeChartAxisTickMark.Cross : OfficeChartAxisTickMark.None;
    }
}
