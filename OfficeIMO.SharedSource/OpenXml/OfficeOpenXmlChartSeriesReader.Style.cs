using static OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTextReader;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        /// <summary>
        /// Reads the native chart appearance. In best-effort projections a chart whose surfaces or text
        /// cannot be read keeps <paramref name="textStyle"/> (or the default style) and reports why.
        /// </summary>
        internal static OfficeChartStyle ReadStyleOrApproximate(DocumentFormat.OpenXml.Packaging.ChartPart chartPart, C.Chart chart,
            OfficeChartType kind, A.ColorScheme? scheme, OfficeChartStyle? textStyle = null, int radialPointCount = 0) {
            int checkpoint = OfficeOpenXmlChartFidelityScope.Checkpoint;
            try {
                return ReadStyle(chartPart, chart, kind, scheme, textStyle, radialPointCount);
            } catch (NotSupportedException exception) when (OfficeOpenXmlChartFidelityScope.AllowsApproximation) {
                OfficeOpenXmlChartFidelityScope.RollBack(checkpoint);
                OfficeOpenXmlChartFidelityScope.Record("The chart surfaces and text are drawn with default formatting: " + exception.Message);
                return textStyle ?? OfficeChartStyle.Default;
            }
        }

        internal static OfficeChartStyle ReadStyle(DocumentFormat.OpenXml.Packaging.ChartPart chartPart, C.Chart chart, OfficeChartType kind, A.ColorScheme? scheme, OfficeChartStyle? textStyle = null,
            int radialPointCount = 0) {
            OpenXmlElement? colorMap = OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(chartPart);
            if (chart.Parent?.GetFirstChild<C.RoundedCorners>()?.Val?.Value == true)
                OfficeOpenXmlChartFidelityScope.Approximate("Rounded native chart frames cannot be projected.",
                    "The rounded chart frame is drawn with square corners.");
            if (chart.GetFirstChild<C.Legend>()?.Elements<C.LegendEntry>().Any(entry => entry.GetFirstChild<C.TextProperties>() != null) == true)
                OfficeOpenXmlChartFidelityScope.Approximate("Per-entry legend text formatting cannot be projected.",
                    "Legend entries with their own text formatting use the legend's text formatting.");
            foreach (C.Title title in chart.Descendants<C.Title>()) {
                if (title.GetFirstChild<C.ChartShapeProperties>() is not C.ChartShapeProperties titleShape) continue;
                if (!titleShape.HasChildren && !titleShape.HasAttributes) continue;
                Surface titleSurface = ReadSurface(titleShape, scheme, colorMap);
                if (!titleSurface.NoFill || !titleSurface.NoOutline || titleSurface.Fill.HasValue || titleSurface.Stroke.HasValue)
                    OfficeOpenXmlChartFidelityScope.Approximate("The title shape appearance cannot be projected.",
                        "A title's fill or border is left out.");
            }
            var legendShape = chart.GetFirstChild<C.Legend>()?.GetFirstChild<C.ChartShapeProperties>();
            var legendSurface = ReadSurface(legendShape, scheme, colorMap);
            if (legendSurface.Dash is OfficeStrokeDashStyle legendDash && legendDash != OfficeStrokeDashStyle.Solid)
                OfficeOpenXmlChartFidelityScope.Approximate("The legend border dash cannot be projected.",
                    "The dashed legend border is drawn solid.");
            bool allowMultiline = kind == OfficeChartType.Bubble;
            var titleText = ReadNativeText(chart, chart.GetFirstChild<C.Title>(), scheme, allowMultiline);
            var defaultText = ReadNativeText(chart, chart, scheme, allowMultiline);
            var labelText = kind == OfficeChartType.Bubble ? new LabelLayout() : ReadLabels(chart, scheme);
            var legendText = ReadNativeText(chart, chart.GetFirstChild<C.Legend>(), scheme, allowMultiline);
            var axisText = ReadUniformNativeText(chart, TextAxes(chart), scheme, allowMultiline);
            var axisTitleText = ReadUniformNativeText(chart, TextAxes(chart).Select(axis => axis.GetFirstChild<C.Title>()).Where(title => title != null).Cast<OpenXmlElement>(), scheme, allowMultiline);
            var area = ReadSurface(chart.Parent?.GetFirstChild<C.ShapeProperties>(), scheme, colorMap);
            var plot = chart.PlotArea;
            var plotStyle = ReadSurface(plot?.GetFirstChild<C.ShapeProperties>(), scheme, colorMap);
            OpenXmlCompositeElement? categoryAxis = null, valueAxis = null;
            if (plot != null && kind != OfficeChartType.Pie && kind != OfficeChartType.Doughnut) {
                var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
                var layer = plot.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(element =>
                    element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Primary);
                var axes = layer?.Elements<C.AxisId>().Select(reference => groups.Resolve(reference.Val?.Value)).ToArray();
                bool numeric = kind == OfficeChartType.Scatter || kind == OfficeChartType.Bubble;
                categoryAxis = numeric ? axes?.FirstOrDefault() : axes?.FirstOrDefault(axis => axis is C.CategoryAxis || axis is C.DateAxis);
                valueAxis = numeric ? axes?.Skip(1).FirstOrDefault() : axes?.FirstOrDefault(axis => axis is C.ValueAxis);
            }
            var category = ReadSurface(categoryAxis?.GetFirstChild<C.ChartShapeProperties>(), scheme, colorMap);
            var value = ReadSurface(valueAxis?.GetFirstChild<C.ChartShapeProperties>(), scheme, colorMap);
            var categoryMajor = categoryAxis?.GetFirstChild<C.MajorGridlines>();
            var valueMajor = valueAxis?.GetFirstChild<C.MajorGridlines>();
            var categoryMinor = categoryAxis?.GetFirstChild<C.MinorGridlines>();
            var valueMinor = valueAxis?.GetFirstChild<C.MinorGridlines>();
            var categoryGrid = ReadSurface(categoryMajor?.GetFirstChild<C.ChartShapeProperties>(), scheme, colorMap);
            var valueGrid = ReadSurface(valueMajor?.GetFirstChild<C.ChartShapeProperties>(), scheme, colorMap);
            var categoryMinorGrid = ReadSurface(categoryMinor?.GetFirstChild<C.ChartShapeProperties>(), scheme, colorMap);
            var valueMinorGrid = ReadSurface(valueMinor?.GetFirstChild<C.ChartShapeProperties>(), scheme, colorMap);
            C.PieChartSeries? radialSeries = kind is OfficeChartType.Pie or OfficeChartType.Doughnut
                ? plot?.Descendants<C.PieChartSeries>().FirstOrDefault() : null;
            OfficeColor[]? radialPalette = radialSeries == null ? null :
                OfficeOpenXmlThemeColorResolver.ReadRadialPalette(chartPart, radialSeries, radialPointCount, scheme);
            return new OfficeChartStyle(palette: radialPalette,
                fontFamily: textStyle?.FontFamily, titleFontFamily: textStyle?.TitleFontFamily,
                titleFontSize: titleText.Size, titleFontStyle: titleText.Style, titleColor: titleText.Color,
                legendTextColor: legendText.Color, mutedTextColor: axisText.Color, axisTitleColor: axisTitleText.Color,
                legendBackgroundColor: legendSurface.NoFill ? null : legendSurface.Fill,
                legendBorderColor: legendSurface.NoOutline ? null : legendSurface.Stroke,
                legendBorderWidth: legendSurface.Width,
                textColor: defaultText.Color, dataLabelTextColor: labelText.TextColor ?? defaultText.Color,
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

        private static Surface ReadSurface(OpenXmlElement? properties, A.ColorScheme? scheme, OpenXmlElement? colorMap) {
            if (properties == null) return default;
            // An empty effect list, as Office writes on most chart surfaces, draws nothing.
            if (properties.ChildElements.Any(child => child is not A.SolidFill && child is not A.NoFill && child is not A.Outline &&
                !(child is A.EffectList && !child.HasChildren)))
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface has an unsupported fill or effect.",
                    "A chart surface fill or effect the renderer cannot draw is left out.");
            var outline = properties.GetFirstChild<A.Outline>();
            // Office writes cap="flat" cmpd="sng" algn="ctr" and a round join on every chart surface.
            // Those defaults draw exactly; caps and joins only touch frame corners and axis-line ends,
            // which stay within a fraction of the hairline width at chart scale.
            if (outline?.CompoundLineType?.Value is A.CompoundLineValues compound && compound != A.CompoundLineValues.Single)
                OfficeOpenXmlChartFidelityScope.Approximate("The chart outline attributes cannot be projected.",
                    "A compound chart outline is drawn as a single line.");
            if (outline?.Alignment?.Value is A.PenAlignmentValues alignment && alignment != A.PenAlignmentValues.Center)
                OfficeOpenXmlChartFidelityScope.Approximate("The chart outline attributes cannot be projected.",
                    "An inset chart outline is drawn centred on its edge.");
            if (outline?.ChildElements.Any(child => child is not A.SolidFill && child is not A.NoFill && child is not A.PresetDash &&
                    child is not A.Round && child is not A.LineJoinBevel && !(child is A.Miter miter && !miter.HasChildren)) == true)
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface has an unsupported outline.",
                    "A chart outline fill or line end the renderer cannot draw is drawn as a plain line.");
            if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(properties.GetFirstChild<A.SolidFill>()) ||
                OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(outline?.GetFirstChild<A.SolidFill>()))
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface has an unsupported colour transform.",
                    "A chart surface colour transform the renderer cannot apply is ignored.");
            OfficeColor? fill = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.SolidFill>(), scheme, colorMap: colorMap);
            OfficeColor? stroke = OfficeOpenXmlThemeColorResolver.ResolveColor(outline?.GetFirstChild<A.SolidFill>(), scheme, colorMap: colorMap);
            if (properties.GetFirstChild<A.SolidFill>() != null && !fill.HasValue || outline?.GetFirstChild<A.SolidFill>() != null && !stroke.HasValue)
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface colour cannot be resolved.",
                    "A chart surface colour that cannot be resolved is drawn with the default chart colour.");
            if (outline != null && outline.GetFirstChild<A.NoFill>() == null && !stroke.HasValue)
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface outline colour cannot be projected.",
                    "A chart surface outline colour that cannot be resolved is drawn with the default chart colour.");
            double? width = outline?.Width?.Value is int emus ? emus / 12700d : null;
            if (width.HasValue && (width < 0 || width > OfficeChartStyleBounds.MaximumLineWidthPoints))
                throw new NotSupportedException("The chart outline width is outside the supported range.");
            if (width == 0) {
                width = outline?.GetFirstChild<A.NoFill>() == null
                    ? OfficeChartStyleBounds.HairlineWidthPoints
                    : null;
            }
            var dash = ReadDash(outline);
            if (outline?.GetFirstChild<A.PresetDash>() != null && !dash.HasValue)
                throw new NotSupportedException("The chart outline dash pattern cannot be projected.");
            return new Surface(fill, stroke, width, dash, properties.GetFirstChild<A.NoFill>() != null, outline?.GetFirstChild<A.NoFill>() != null);
        }
    }
}
