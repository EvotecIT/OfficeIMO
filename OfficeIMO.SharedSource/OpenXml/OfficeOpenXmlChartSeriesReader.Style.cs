using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static OfficeChartStyle ReadStyle(C.Chart chart, OfficeChartKind kind, A.ColorScheme? scheme, OfficeChartStyle? textStyle = null) {
            if (chart.Parent?.GetFirstChild<C.RoundedCorners>()?.Val?.Value == true)
                throw new NotSupportedException("Rounded native chart frames cannot be projected.");
            if (chart.GetFirstChild<C.Legend>()?.Elements<C.LegendEntry>().Any(entry => entry.GetFirstChild<C.TextProperties>() != null) == true)
                throw new NotSupportedException("Per-entry legend text formatting cannot be projected.");
            if (chart.Descendants<C.Title>().Cast<OpenXmlElement>().Concat(chart.Elements<C.Legend>())
                .Any(owner => owner?.GetFirstChild<C.ChartShapeProperties>()?.ChildElements.Count > 0))
                throw new NotSupportedException("Title and legend shape appearance cannot be projected.");
            var titleText = ReadNativeText(chart, chart.GetFirstChild<C.Title>(), scheme);
            var defaultText = ReadNativeText(chart, chart, scheme);
            var legendText = ReadNativeText(chart, chart.GetFirstChild<C.Legend>(), scheme);
            var axisText = ReadUniformNativeText(chart, TextAxes(chart), scheme);
            var axisTitleText = ReadUniformNativeText(chart, TextAxes(chart).Select(axis => axis.GetFirstChild<C.Title>()).Where(title => title != null).Cast<OpenXmlElement>(), scheme);
            var area = ReadSurface(chart.Parent?.GetFirstChild<C.ShapeProperties>(), scheme);
            var plot = chart.PlotArea;
            var plotStyle = ReadSurface(plot?.GetFirstChild<C.ShapeProperties>(), scheme);
            OpenXmlCompositeElement? categoryAxis = null, valueAxis = null;
            if (plot != null && kind != OfficeChartKind.Pie && kind != OfficeChartKind.Doughnut) {
                var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
                var layer = plot.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(element =>
                    element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Primary);
                var axes = layer?.Elements<C.AxisId>().Select(reference => groups.Resolve(reference.Val?.Value)).ToArray();
                bool numeric = kind == OfficeChartKind.Scatter || kind == OfficeChartKind.Bubble;
                categoryAxis = numeric ? axes?.FirstOrDefault() : axes?.FirstOrDefault(axis => axis is C.CategoryAxis || axis is C.DateAxis);
                valueAxis = numeric ? axes?.Skip(1).FirstOrDefault() : axes?.FirstOrDefault(axis => axis is C.ValueAxis);
            }
            var category = ReadSurface(categoryAxis?.GetFirstChild<C.ChartShapeProperties>(), scheme);
            var value = ReadSurface(valueAxis?.GetFirstChild<C.ChartShapeProperties>(), scheme);
            var categoryMajor = categoryAxis?.GetFirstChild<C.MajorGridlines>();
            var valueMajor = valueAxis?.GetFirstChild<C.MajorGridlines>();
            var categoryMinor = categoryAxis?.GetFirstChild<C.MinorGridlines>();
            var valueMinor = valueAxis?.GetFirstChild<C.MinorGridlines>();
            var categoryGrid = ReadSurface(categoryMajor?.GetFirstChild<C.ChartShapeProperties>(), scheme);
            var valueGrid = ReadSurface(valueMajor?.GetFirstChild<C.ChartShapeProperties>(), scheme);
            var categoryMinorGrid = ReadSurface(categoryMinor?.GetFirstChild<C.ChartShapeProperties>(), scheme);
            var valueMinorGrid = ReadSurface(valueMinor?.GetFirstChild<C.ChartShapeProperties>(), scheme);
            return new OfficeChartStyle(fontFamily: textStyle?.FontFamily, titleFontFamily: textStyle?.TitleFontFamily,
                titleFontSize: titleText.Size, titleFontStyle: titleText.Style, titleColor: titleText.Color,
                legendTextColor: legendText.Color, mutedTextColor: axisText.Color, axisTitleColor: axisTitleText.Color,
                textColor: defaultText.Color, dataLabelTextColor: defaultText.Color,
                showBackground: !area.NoFill, backgroundColor: area.Fill,
                showBorder: !area.NoOutline, borderColor: area.Stroke, chartBorderWidth: area.Width, chartBorderDashStyle: area.Dash,
                plotAreaBackgroundColor: plotStyle.NoFill ? null : plotStyle.Fill,
                plotAreaBorderColor: plotStyle.NoOutline ? null : plotStyle.Stroke, plotAreaBorderWidth: plotStyle.Width, plotAreaBorderDashStyle: plotStyle.Dash,
                categoryAxisColor: category.NoOutline ? OfficeColor.Transparent : category.Stroke,
                valueAxisColor: value.NoOutline ? OfficeColor.Transparent : value.Stroke,
                categoryAxisLineWidth: category.Width, valueAxisLineWidth: value.Width,
                categoryAxisLineDashStyle: category.Dash, valueAxisLineDashStyle: value.Dash,
                categoryGridLineColor: categoryGrid.Stroke, valueGridLineColor: valueGrid.Stroke,
                categoryGridLineWidth: categoryGrid.Width, valueGridLineWidth: valueGrid.Width,
                categoryGridLineDashStyle: categoryGrid.Dash, valueGridLineDashStyle: valueGrid.Dash,
                showCategoryGridLines: categoryMajor != null && !categoryGrid.NoOutline,
                showValueGridLines: valueMajor != null && !valueGrid.NoOutline,
                categoryMinorGridLineColor: categoryMinorGrid.Stroke, valueMinorGridLineColor: valueMinorGrid.Stroke,
                categoryMinorGridLineWidth: categoryMinorGrid.Width, valueMinorGridLineWidth: valueMinorGrid.Width,
                categoryMinorGridLineDashStyle: categoryMinorGrid.Dash, valueMinorGridLineDashStyle: valueMinorGrid.Dash,
                showCategoryMinorGridLines: categoryMinor != null && !categoryMinorGrid.NoOutline,
                showValueMinorGridLines: valueMinor != null && !valueMinorGrid.NoOutline);
        }

        private readonly struct Surface {
            internal Surface(OfficeColor? fill, OfficeColor? stroke, double? width, OfficeStrokeDashStyle? dash, bool noFill, bool noOutline) {
                Fill = fill; Stroke = stroke; Width = width; Dash = dash; NoFill = noFill; NoOutline = noOutline;
            }
            internal OfficeColor? Fill { get; }
            internal OfficeColor? Stroke { get; }
            internal double? Width { get; }
            internal OfficeStrokeDashStyle? Dash { get; }
            internal bool NoFill { get; }
            internal bool NoOutline { get; }
        }

        private static Surface ReadSurface(OpenXmlElement? properties, A.ColorScheme? scheme) {
            if (properties == null) return default;
            if (properties.ChildElements.Any(child => child is not A.SolidFill && child is not A.NoFill && child is not A.Outline))
                throw new NotSupportedException("The chart surface has an unsupported fill or effect.");
            var outline = properties.GetFirstChild<A.Outline>();
            if (outline?.CapType != null || outline?.Alignment != null || outline?.CompoundLineType != null)
                throw new NotSupportedException("The chart outline attributes cannot be projected.");
            if (outline?.ChildElements.Any(child => child is not A.SolidFill && child is not A.NoFill && child is not A.PresetDash) == true)
                throw new NotSupportedException("The chart surface has an unsupported outline.");
            if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(properties.GetFirstChild<A.SolidFill>()) ||
                OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(outline?.GetFirstChild<A.SolidFill>()))
                throw new NotSupportedException("The chart surface has an unsupported colour transform.");
            OfficeColor? fill = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.SolidFill>(), scheme);
            OfficeColor? stroke = OfficeOpenXmlThemeColorResolver.ResolveColor(outline?.GetFirstChild<A.SolidFill>(), scheme);
            if (properties.GetFirstChild<A.SolidFill>() != null && !fill.HasValue || outline?.GetFirstChild<A.SolidFill>() != null && !stroke.HasValue)
                throw new NotSupportedException("The chart surface colour cannot be resolved.");
            double? width = outline?.Width?.Value is int emus ? emus / 12700d : null;
            if (width.HasValue && (width <= 0 || width > OfficeChartStyleBounds.MaximumLineWidthPoints))
                throw new NotSupportedException("The chart outline width is outside the supported range.");
            var dash = ReadDash(outline);
            if (outline?.GetFirstChild<A.PresetDash>() != null && !dash.HasValue)
                throw new NotSupportedException("The chart outline dash pattern cannot be projected.");
            return new Surface(fill, stroke, width, dash, properties.GetFirstChild<A.NoFill>() != null, outline?.GetFirstChild<A.NoFill>() != null);
        }
    }
}
