using static OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTextReader;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        /// <summary>
        /// Reads the native chart appearance. In best-effort projections a chart whose surfaces or text
        /// cannot be read keeps <paramref name="textStyle"/> (or the default style) and reports why.
        /// </summary>
        internal static OfficeChartStyle ReadStyleOrApproximate(IOfficeChartPart chartPart, OfficeXmlNode chart,
            OfficeChartType kind, OfficeXmlNode scheme, OfficeChartStyle? textStyle = null, int radialPointCount = 0) {
            int checkpoint = OfficeOpenXmlChartFidelityScope.Checkpoint;
            try {
                return ReadStyle(chartPart, chart, kind, scheme, textStyle, radialPointCount);
            } catch (NotSupportedException exception) when (OfficeOpenXmlChartFidelityScope.AllowsApproximation) {
                OfficeOpenXmlChartFidelityScope.RollBack(checkpoint);
                OfficeOpenXmlChartFidelityScope.Record("The chart surfaces and text are drawn with default formatting: " + exception.Message);
                return textStyle ?? OfficeChartStyle.Default;
            }
        }

        internal static OfficeChartStyle ReadStyle(IOfficeChartPart chartPart, OfficeXmlNode chart, OfficeChartType kind, OfficeXmlNode scheme, OfficeChartStyle? textStyle = null,
            int radialPointCount = 0) {
            OfficeXmlNode colorMap = OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(chartPart);
            if (chart.Parent.Child(NsC, "roundedCorners").BoolVal() == true)
                OfficeOpenXmlChartFidelityScope.Approximate("Rounded native chart frames cannot be projected.",
                    "The rounded chart frame is drawn with square corners.");
            if (chart.Child(NsC, "legend").Children(NsC, "legendEntry").ToList().Any(entry => !entry.Child(NsC, "txPr").IsNone))
                OfficeOpenXmlChartFidelityScope.Approximate("Per-entry legend text formatting cannot be projected.",
                    "Legend entries with their own text formatting use the legend's text formatting.");
            foreach (OfficeXmlNode title in chart.Descendants(NsC, "title")) {
                OfficeXmlNode titleShape = title.Child(NsC, "spPr");
                if (titleShape.IsNone) continue;
                if (!titleShape.HasChildElements() && !titleShape.HasAttributes) continue;
                Surface titleSurface = ReadSurface(titleShape, scheme, colorMap);
                if (!titleSurface.NoFill || !titleSurface.NoOutline || titleSurface.Fill.HasValue || titleSurface.Stroke.HasValue)
                    OfficeOpenXmlChartFidelityScope.Approximate("The title shape appearance cannot be projected.",
                        "A title's fill or border is left out.");
            }
            var legendShape = chart.Child(NsC, "legend").Child(NsC, "spPr");
            var legendSurface = ReadSurface(legendShape, scheme, colorMap);
            if (legendSurface.Dash is OfficeStrokeDashStyle legendDash && legendDash != OfficeStrokeDashStyle.Solid)
                OfficeOpenXmlChartFidelityScope.Approximate("The legend border dash cannot be projected.",
                    "The dashed legend border is drawn solid.");
            bool allowMultiline = kind == OfficeChartType.Bubble;
            var titleText = ReadNativeText(chart, chart.Child(NsC, "title"), scheme, allowMultiline);
            var defaultText = ReadNativeText(chart, chart, scheme, allowMultiline);
            var labelText = kind == OfficeChartType.Bubble ? new LabelLayout() : ReadLabels(chart, scheme);
            var legendText = ReadNativeText(chart, chart.Child(NsC, "legend"), scheme, allowMultiline);
            var axisText = ReadUniformNativeText(chart, TextAxes(chart), scheme, allowMultiline);
            var axisTitleText = ReadUniformNativeText(chart, TextAxes(chart).Select(axis => axis.Child(NsC, "title")).Where(title => !title.IsNone), scheme, allowMultiline);
            var area = ReadSurface(chart.Parent.Child(NsC, "spPr"), scheme, colorMap);
            var plot = chart.Child(NsC, "plotArea");
            var plotStyle = ReadSurface(plot.Child(NsC, "spPr"), scheme, colorMap);
            OfficeXmlNode categoryAxis = OfficeXmlNode.None, valueAxis = OfficeXmlNode.None;
            if (!plot.IsNone && kind != OfficeChartType.Pie && kind != OfficeChartType.Doughnut) {
                var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
                var layer = plot.Children().ToList().FirstOrDefault(element =>
                    element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && groups.Read(element) == OfficeChartAxisGroup.Primary);
                var axes = layer.IsNone ? null : layer.Children(NsC, "axId").ToList().Select(reference => groups.Resolve(reference.UIntVal())).ToArray();
                bool numeric = kind == OfficeChartType.Scatter || kind == OfficeChartType.Bubble;
                categoryAxis = (numeric ? axes?.FirstOrDefault() : axes?.FirstOrDefault(axis => axis.IsChart("catAx") || axis.IsChart("dateAx"))) ?? OfficeXmlNode.None;
                valueAxis = (numeric ? axes?.Skip(1).FirstOrDefault() : axes?.FirstOrDefault(axis => axis.IsChart("valAx"))) ?? OfficeXmlNode.None;
            }
            var category = ReadSurface(categoryAxis.Child(NsC, "spPr"), scheme, colorMap);
            var value = ReadSurface(valueAxis.Child(NsC, "spPr"), scheme, colorMap);
            var categoryMajor = categoryAxis.Child(NsC, "majorGridlines");
            var valueMajor = valueAxis.Child(NsC, "majorGridlines");
            var categoryMinor = categoryAxis.Child(NsC, "minorGridlines");
            var valueMinor = valueAxis.Child(NsC, "minorGridlines");
            var categoryGrid = ReadSurface(categoryMajor.Child(NsC, "spPr"), scheme, colorMap);
            var valueGrid = ReadSurface(valueMajor.Child(NsC, "spPr"), scheme, colorMap);
            var categoryMinorGrid = ReadSurface(categoryMinor.Child(NsC, "spPr"), scheme, colorMap);
            var valueMinorGrid = ReadSurface(valueMinor.Child(NsC, "spPr"), scheme, colorMap);
            OfficeXmlNode radialSeries = kind is OfficeChartType.Pie or OfficeChartType.Doughnut
                ? plot.Descendants(NsC, "ser").FirstOrDefault(series => IsPieSeriesGroup(series.Parent))
                : OfficeOpenXmlChartPlotSemantics.GetPointColoredSeries(plot);
            OfficeColor[]? radialPalette = radialSeries.IsNone ? null :
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
                showCategoryGridLines: !categoryMajor.IsNone && !categoryGrid.NoOutline,
                showValueGridLines: !valueMajor.IsNone && !valueGrid.NoOutline,
                categoryMinorGridLineColor: categoryMinorGrid.Stroke, valueMinorGridLineColor: valueMinorGrid.Stroke,
                categoryMinorGridLineWidth: categoryMinorGrid.Width, valueMinorGridLineWidth: valueMinorGrid.Width,
                categoryMinorGridLineDashStyle: categoryMinorGrid.Dash, valueMinorGridLineDashStyle: valueMinorGrid.Dash,
                showCategoryMinorGridLines: !categoryMinor.IsNone && !categoryMinorGrid.NoOutline,
                showValueMinorGridLines: !valueMinor.IsNone && !valueMinorGrid.NoOutline);
        }

        // The chart groups whose c:ser the SDK types as a pie series.
        private static bool IsPieSeriesGroup(OfficeXmlNode group) =>
            group.NamespaceId == NsC && group.LocalName is "pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart";

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

        private static Surface ReadSurface(OfficeXmlNode properties, OfficeXmlNode scheme, OfficeXmlNode colorMap) {
            if (properties.IsNone) return default;
            // An empty effect list, as Office writes on most chart surfaces, draws nothing.
            if (properties.Children().ToList().Any(child => !child.IsDrawing("solidFill") && !child.IsDrawing("noFill") && !child.IsDrawing("ln") &&
                !(child.IsDrawing("effectLst") && !child.HasChildElements())))
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface has an unsupported fill or effect.",
                    "A chart surface fill or effect the renderer cannot draw is left out.");
            var outline = properties.Child(NsA, "ln");
            // Office writes cap="flat" cmpd="sng" algn="ctr" and a round join on every chart surface.
            // Those defaults draw exactly; caps and joins only touch frame corners and axis-line ends,
            // which stay within a fraction of the hairline width at chart scale.
            if (outline.Attribute("cmpd") is string compound && compound != "sng")
                OfficeOpenXmlChartFidelityScope.Approximate("The chart outline attributes cannot be projected.",
                    "A compound chart outline is drawn as a single line.");
            if (outline.Attribute("algn") is string alignment && alignment != "ctr")
                OfficeOpenXmlChartFidelityScope.Approximate("The chart outline attributes cannot be projected.",
                    "An inset chart outline is drawn centred on its edge.");
            if (outline.Children().ToList().Any(child => !child.IsDrawing("solidFill") && !child.IsDrawing("noFill") && !child.IsDrawing("prstDash") &&
                    !child.IsDrawing("round") && !child.IsDrawing("bevel") && !(child.IsDrawing("miter") && !child.HasChildElements())))
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface has an unsupported outline.",
                    "A chart outline fill or line end the renderer cannot draw is drawn as a plain line.");
            if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(properties.Child(NsA, "solidFill")) ||
                OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(outline.Child(NsA, "solidFill")))
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface has an unsupported colour transform.",
                    "A chart surface colour transform the renderer cannot apply is ignored.");
            OfficeColor? fill = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.Child(NsA, "solidFill"), scheme, colorMap: colorMap);
            OfficeColor? stroke = OfficeOpenXmlThemeColorResolver.ResolveColor(outline.Child(NsA, "solidFill"), scheme, colorMap: colorMap);
            if (!properties.Child(NsA, "solidFill").IsNone && !fill.HasValue || !outline.Child(NsA, "solidFill").IsNone && !stroke.HasValue)
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface colour cannot be resolved.",
                    "A chart surface colour that cannot be resolved is drawn with the default chart colour.");
            if (!outline.IsNone && outline.Child(NsA, "noFill").IsNone && !stroke.HasValue)
                OfficeOpenXmlChartFidelityScope.Approximate("The chart surface outline colour cannot be projected.",
                    "A chart surface outline colour that cannot be resolved is drawn with the default chart colour.");
            double? width = outline.IntAttr("w") is int emus ? emus / 12700d : null;
            if (width.HasValue && (width < 0 || width > OfficeChartStyleBounds.MaximumLineWidthPoints))
                throw new NotSupportedException("The chart outline width is outside the supported range.");
            if (width == 0) {
                width = outline.Child(NsA, "noFill").IsNone
                    ? OfficeChartStyleBounds.HairlineWidthPoints
                    : null;
            }
            var dash = ReadDash(outline);
            if (!outline.Child(NsA, "prstDash").IsNone && !dash.HasValue)
                throw new NotSupportedException("The chart outline dash pattern cannot be projected.");
            return new Surface(fill, stroke, width, dash, !properties.Child(NsA, "noFill").IsNone, !outline.Child(NsA, "noFill").IsNone);
        }

        // SDK callers
        internal static OfficeChartStyle ReadStyleOrApproximate(DocumentFormat.OpenXml.Packaging.ChartPart chartPart, C.Chart chart,
            OfficeChartType kind, A.ColorScheme? scheme, OfficeChartStyle? textStyle = null, int radialPointCount = 0) =>
            ReadStyleOrApproximate(new OfficeSdkChartPart(chartPart), OfficeXmlNode.FromSdk(chart), kind, OfficeXmlNode.FromSdk(scheme),
                textStyle, radialPointCount);
    }
}
