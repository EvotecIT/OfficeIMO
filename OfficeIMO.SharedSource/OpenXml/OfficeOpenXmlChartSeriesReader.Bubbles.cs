using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static Result? ReadBubbles(IEnumerable<OfficeXmlNode> elements,
            OfficeXmlNode colorScheme, int maximumPoints, bool forDataUpdate = false,
            bool validatePlot = true, ProjectionBudget? projectionBudget = null, int maximumPointOverrides = 1_000_000,
            SourceVisibility? visibility = null) {
            var source = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(elements, maximumPoints);
            if (source.Count == 0) return null;
            if (validatePlot) ValidateSeriesPlotBudget(source[0], maximumPoints);
            projectionBudget ??= new ProjectionBudget();
            if (!forDataUpdate) {
                var orders = new HashSet<uint>();
                foreach (var item in source) {
                    uint? order = item.Child(NsC, "order").UIntVal();
                    if (!order.HasValue || !orders.Add(order.Value)) return null;
                }
                source = source.OrderBy(item => item.Child(NsC, "order").UIntVal()!.Value).ToList();
            }
            var series = new List<Series>();
            IReadOnlyList<string>? categories = null;
            for (int index = 0; index < source.Count; index++) {
                var element = source[index];
                if (!forDataUpdate && !element.Child(NsC, "trendline").IsNone)
                    OfficeOpenXmlChartFidelityScope.Approximate("Trendlines on bubble charts cannot be projected.",
                        "Trendlines on bubble charts are left out.");
                if (!forDataUpdate && (
                    HasUnsupportedSeriesStyle(element.Child(NsC, "spPr")) || HasUnresolvedSeriesColor(element.Child(NsC, "spPr"), colorScheme) ||
                    OfficeOpenXmlChartPointStyles.GetBoundedPoints(element, maximumPointOverrides).Any(point =>
                        HasUnsupportedPointStyle(point.Child(NsC, "spPr")) || HasUnresolvedPointColor(point.Child(NsC, "spPr"), colorScheme)))) return null;
                OfficeXmlNode xSource = element.Child(NsC, "xVal");
                OfficeXmlNode ySource = element.Child(NsC, "yVal");
                OfficeXmlNode sizeSource = element.Child(NsC, "bubbleSize");
                if (!TryReadStrictCachedNumbers(xSource, true, maximumPoints, out var x) ||
                    !TryReadStrictCachedNumbers(ySource, true, maximumPoints, out var y) ||
                    !TryReadStrictCachedNumbers(sizeSource, false, maximumPoints, out var sizes)) return null;
                x = SourceVisibility.Select(visibility, xSource, x);
                y = SourceVisibility.Select(visibility, ySource, y);
                sizes = SourceVisibility.Select(visibility, sizeSource, sizes);
                if (x.Count != y.Count || x.Count != sizes.Count) return null;
                if (!forDataUpdate && y.Any(value => value < 0) && element.Child(NsC, "invertIfNegative") is { IsNone: false } invert && invert.BoolVal() != false) return null;
                if (x.Count == 0) continue;
                projectionBudget.Reserve(x.Count);
                categories ??= x.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray();
                var text = element.Child(NsC, "tx");
                string name = text.ChildText(NsC, "v") ??
                    OfficeOpenXmlChartCacheReader.ReadCachedStrings(text, maximumPoints).FirstOrDefault() ??
                    string.Concat(text.Descendants(NsA, "t").Select(value => value.InnerText));
                if (string.IsNullOrWhiteSpace(name)) {
                    if (!forDataUpdate && !text.Child(NsC, "strRef").IsNone) return null;
                    name = "Series " + (index + 1).ToString(CultureInfo.InvariantCulture);
                }
                var properties = element.Child(NsC, "spPr");
                var outline = properties.Child(NsA, "ln");
                var fill = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.Child(NsA, "solidFill"), colorScheme);
                var stroke = OfficeOpenXmlThemeColorResolver.ResolveColor(outline.Child(NsA, "solidFill"), colorScheme);
                OfficeLength? width = outline.IntAttr("w") is int emus && emus > 0 ? OfficeLength.Emus(emus) : null;
                bool automaticColor = OfficeOpenXmlChartAutomaticColors.HasAutomaticSeriesColor(element);
                if (!forDataUpdate && fill == null) {
                    // Office colours bubbles without a stored fill from the chart style and theme.
                    fill = OfficeOpenXmlChartAutomaticColors.ResolveSeries(element, colorScheme,
                        projectionBudget.AutomaticColors);
                    if (fill == null) return null;
                    automaticColor = true;
                }
                var data = OfficeSeries.CreateBubble(name, x, y, sizes, fill,
                    ReadBubblePointColors(element, x.Count, colorScheme, maximumPointOverrides), markerOutlineColor: stroke,
                    markerOutlineWidth: width, showMarkerOutline: outline.Child(NsA, "noFill").IsNone)
                    .WithSeriesStyle(OfficeOpenXmlChartPointStyles.ReadSeriesStyle(element, colorScheme))
                    .WithPointStyles(forDataUpdate ? null : OfficeOpenXmlChartPointStyles.Read(
                        OfficeOpenXmlChartPointStyles.GetBoundedPoints(element, maximumPointOverrides), x.Count, colorScheme));
                if (!forDataUpdate)
                    data = data.WithErrorBars(OfficeOpenXmlChartErrorBars.Read(element, colorScheme, maximumPoints,
                        visibility == null ? null : (source, amounts) => SourceVisibility.Select(visibility, source, amounts)));
                series.Add(new Series(element.Child(NsC, "idx").UIntVal() ?? (uint)index,
                    data, hasAutomaticColor: automaticColor));
            }
            return categories == null || series.Count == 0 ? null : new Result(categories, series);
        }

        private static bool HasUnsupportedEffects(OfficeXmlNode properties) =>
            properties.Child(NsA, "effectLst").HasChildElements() ||
            properties.Child(NsA, "effectDag").HasChildElements() ||
            properties.Children().ToList().Any(child =>
                child.LocalName == "scene3d" ||
                child.LocalName == "sp3d");

        private static bool HasUnsupportedSeriesStyle(
            OfficeXmlNode properties) =>
            HasUnsupportedPointStyle(properties);

        private static bool HasUnsupportedPointStyle(
            OfficeXmlNode properties) =>
            properties.Children().ToList().Any(child => child.NamespaceId == NsA && child.LocalName is "gradFill" or "blipFill" or "grpFill") ||
            (properties.Child(NsA, "pattFill") is { IsNone: false } pattern &&
                !OfficeOpenXmlChartPointStyles.ReadHatch(pattern.Attribute("prst")).HasValue) ||
            HasUnsupportedEffects(properties) ||
            HasUnsupportedOutlineFill(properties.Child(NsA, "ln"));

        private static bool HasUnsupportedOutlineFill(OfficeXmlNode outline) =>
            !outline.IsNone &&
            (outline.IntAttr("w") == 0 && outline.Child(NsA, "noFill").IsNone ||
             outline.Attribute("algn") is string alignment &&
             alignment != "ctr" ||
             outline.Attribute("cmpd") is string compoundLine &&
             compoundLine != "sng" ||
             outline.Child(NsA, "miter") is { IsNone: false } miter && (miter.HasAttributes || miter.HasChildElements()) ||
             outline.Children().ToList().Any(child =>
                 !(child.NamespaceId == NsA && child.LocalName is "solidFill" or "noFill" or "round" or "bevel" or "miter")));

        private static bool HasUnresolvedSeriesColor(
            OfficeXmlNode properties, OfficeXmlNode colorScheme) {
            if (properties.Child(NsA, "pattFill") is { IsNone: false } pattern)
                return !OfficeOpenXmlChartPointStyles.IsSupportedPattern(pattern, colorScheme) ||
                    HasUnresolvedPointColor(properties, colorScheme);
            if (properties.Child(NsA, "noFill").IsNone &&
                (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(properties.Child(NsA, "solidFill")) ||
                !OfficeOpenXmlThemeColorResolver.ResolveColor(
                    properties.Child(NsA, "solidFill"), colorScheme).HasValue)) {
                return true;
            }

            OfficeXmlNode outline = properties.Child(NsA, "ln");
            if (outline.IsNone ||
                (outline.Child(NsA, "noFill").IsNone &&
                 outline.Child(NsA, "solidFill").IsNone)) {
                return true;
            }

            return HasUnresolvedSolidFill(
                outline.Child(NsA, "solidFill"), colorScheme);
        }

        private static bool HasUnresolvedPointColor(
            OfficeXmlNode properties, OfficeXmlNode colorScheme) {
            if (HasUnresolvedSolidFill(properties.Child(NsA, "solidFill"), colorScheme) ||
                HasUnresolvedSolidFill(properties.Child(NsA, "ln").Child(NsA, "solidFill"), colorScheme)) return true;
            if (properties.Child(NsA, "pattFill") is { IsNone: false } pattern)
                return OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(pattern.Child(NsA, "fgClr")) ||
                    OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(pattern.Child(NsA, "bgClr")) ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.Child(NsA, "fgClr"), colorScheme).HasValue ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.Child(NsA, "bgClr"), colorScheme).HasValue;
            return false;
        }

        private static bool HasUnresolvedSolidFill(
            OfficeXmlNode fill, OfficeXmlNode colorScheme) =>
            !fill.IsNone &&
            (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(fill) ||
             !OfficeOpenXmlThemeColorResolver.ResolveColor(fill, colorScheme).HasValue);

        private static IReadOnlyList<OfficeColor?>? ReadBubblePointColors(
            OfficeXmlNode series, int pointCount, OfficeXmlNode colorScheme, int maximumPointOverrides) {
            var colors = new OfficeColor?[pointCount];
            bool found = false;
            foreach (OfficeXmlNode point in OfficeOpenXmlChartPointStyles.GetBoundedPoints(series, maximumPointOverrides)) {
                uint? sourceIndex = point.Child(NsC, "idx").UIntVal();
                if (!sourceIndex.HasValue || sourceIndex.Value >= (uint)pointCount) continue;
                OfficeXmlNode properties =
                    point.Child(NsC, "spPr");
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                    properties.Child(NsA, "solidFill"), colorScheme);
                if (!color.HasValue) continue;
                colors[(int)sourceIndex.Value] = color;
                found = true;
            }
            return found ? colors : null;
        }

        private static bool TryReadStrictCachedNumbers(OfficeXmlNode container,
            bool allowNegative, int maximumPoints, out IReadOnlyList<double> values) {
            values = Array.Empty<double>();
            if (container.IsNone) {
                return false;
            }

            List<OfficeXmlNode> points =
                OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(CachePoints(container, numeric: true), maximumPoints);
            if (points.Count == 0) {
                return false;
            }

            int length = OfficeOpenXmlChartCacheReader.GetCachedPointLength(container, points,
                point => point.UIntAttr("idx"), maximumPoints);
            if (length != points.Count) {
                return false;
            }

            var parsed = new double[length];
            var seen = new bool[length];
            for (int pointIndex = 0; pointIndex < points.Count; pointIndex++) {
                OfficeXmlNode point = points[pointIndex];
                uint rawIndex = point.UIntAttr("idx") ?? (uint)pointIndex;
                if (rawIndex >= (uint)length || seen[(int)rawIndex]) {
                    return false;
                }

                string? text = point.ChildText(NsC, "v");
                if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                    double.IsNaN(value) || double.IsInfinity(value) ||
                    (!allowNegative && value < 0D)) {
                    return false;
                }

                parsed[(int)rawIndex] = value;
                seen[(int)rawIndex] = true;
            }

            if (seen.Any(present => !present)) {
                return false;
            }

            values = parsed;
            return true;
        }

        // SDK callers
        internal static Result? ReadBubbles(IEnumerable<C.BubbleChartSeries> elements,
            A.ColorScheme? colorScheme, int maximumPoints, bool forDataUpdate = false,
            bool validatePlot = true, ProjectionBudget? projectionBudget = null, int maximumPointOverrides = 1_000_000,
            SourceVisibility? visibility = null) =>
            ReadBubbles(elements.Select(element => OfficeXmlNode.FromSdk(element)), OfficeXmlNode.FromSdk(colorScheme), maximumPoints,
                forDataUpdate, validatePlot, projectionBudget, maximumPointOverrides, visibility);
    }
}
