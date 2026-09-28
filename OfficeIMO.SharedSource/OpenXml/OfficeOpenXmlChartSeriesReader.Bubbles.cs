using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static Result? ReadBubbles(IEnumerable<C.BubbleChartSeries> elements,
            ColorScheme? colorScheme, int maximumPoints, bool forDataUpdate = false,
            bool validatePlot = true, ProjectionBudget? projectionBudget = null) {
            var source = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(elements, maximumPoints);
            if (source.Count == 0) return null;
            if (validatePlot) ValidatePlotBudget(source[0], maximumPoints);
            projectionBudget ??= new ProjectionBudget();
            if (!forDataUpdate) {
                var orders = new HashSet<uint>();
                foreach (var item in source) {
                    uint? order = item.GetFirstChild<C.Order>()?.Val?.Value;
                    if (!order.HasValue || !orders.Add(order.Value)) return null;
                }
                source = source.OrderBy(item => item.GetFirstChild<C.Order>()!.Val!.Value).ToList();
            }
            var series = new List<Series>();
            IReadOnlyList<string>? categories = null;
            for (int index = 0; index < source.Count; index++) {
                var element = source[index];
                if (!forDataUpdate && (element.Elements<C.Trendline>().Any() || element.Elements<C.ErrorBars>().Any() ||
                    HasUnsupportedSeriesStyle(element.ChartShapeProperties) || HasUnresolvedSeriesColor(element.ChartShapeProperties, colorScheme) ||
                    OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(element.Elements<C.DataPoint>(), maximumPoints).Any(point =>
                        HasUnsupportedPointStyle(point.ChartShapeProperties) || HasUnresolvedPointColor(point.ChartShapeProperties, colorScheme)))) return null;
                if (!TryReadStrictCachedNumbers(element.GetFirstChild<C.XValues>(), true, maximumPoints, out var x) ||
                    !TryReadStrictCachedNumbers(element.GetFirstChild<C.YValues>(), true, maximumPoints, out var y) ||
                    !TryReadStrictCachedNumbers(element.GetFirstChild<C.BubbleSize>(), false, maximumPoints, out var sizes) ||
                    x.Count != y.Count || x.Count != sizes.Count) return null;
                if (!forDataUpdate && y.Any(value => value < 0) && element.GetFirstChild<C.InvertIfNegative>() is C.InvertIfNegative invert && invert.Val?.Value != false) return null;
                if (x.Count == 0) continue;
                projectionBudget.Reserve(x.Count);
                categories ??= x.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray();
                var text = element.GetFirstChild<C.SeriesText>();
                string name = text?.GetFirstChild<C.NumericValue>()?.Text ??
                    OfficeOpenXmlChartCacheReader.ReadCachedStrings(text, maximumPoints).FirstOrDefault() ??
                    string.Concat(text?.Descendants<Text>().Select(value => value.Text) ?? Enumerable.Empty<string>());
                if (string.IsNullOrWhiteSpace(name)) {
                    if (!forDataUpdate && text?.GetFirstChild<C.StringReference>() != null) return null;
                    name = "Series " + (index + 1).ToString(CultureInfo.InvariantCulture);
                }
                var properties = element.ChartShapeProperties;
                var outline = properties?.GetFirstChild<Outline>();
                var fill = OfficeOpenXmlThemeColorResolver.ResolveColor(properties?.GetFirstChild<SolidFill>(), colorScheme);
                var stroke = OfficeOpenXmlThemeColorResolver.ResolveColor(outline?.GetFirstChild<SolidFill>(), colorScheme);
                double? width = outline?.Width?.Value is int emus && emus > 0 ? emus / 12700d : null;
                var data = OfficeChartSeries.CreateBubble(name, x, y, sizes, fill,
                    ReadBubblePointColors(element, x.Count, colorScheme, maximumPoints), markerOutlineColor: stroke,
                    markerOutlineWidth: width, showMarkerOutline: outline?.GetFirstChild<NoFill>() == null)
                    .WithPointStyles(forDataUpdate ? null : OfficeOpenXmlChartPointStyles.Read(element, x.Count, colorScheme));
                series.Add(new Series(element.GetFirstChild<C.Index>()?.Val?.Value ?? (uint)index, data));
            }
            return categories == null || series.Count == 0 ? null : new Result(categories, series);
        }

        private static bool HasUnsupportedFill(C.ChartShapeProperties? properties) =>
            properties?.ChildElements.Any(child =>
                child is NoFill or GradientFill or PatternFill or BlipFill or GroupFill) == true;

        private static bool HasUnsupportedEffects(C.ChartShapeProperties? properties) =>
            properties?.GetFirstChild<EffectList>()?.ChildElements.Count > 0 ||
            properties?.GetFirstChild<EffectDag>()?.ChildElements.Count > 0 ||
            properties?.ChildElements.Any(child =>
                child.LocalName == "scene3d" ||
                child.LocalName == "sp3d") == true;

        private static bool HasUnsupportedSeriesStyle(
            C.ChartShapeProperties? properties) =>
            HasUnsupportedFill(properties) ||
            HasUnsupportedEffects(properties) ||
            HasUnsupportedOutlineFill(properties?.GetFirstChild<Outline>());

        private static bool HasUnsupportedPointStyle(
            C.ChartShapeProperties? properties) =>
            properties?.ChildElements.Any(child => child is GradientFill or BlipFill or GroupFill) == true ||
            (properties?.GetFirstChild<PatternFill>() is PatternFill pattern &&
                !OfficeOpenXmlChartPointStyles.ReadHatch(pattern.Preset?.InnerText).HasValue) ||
            HasUnsupportedEffects(properties) ||
            HasUnsupportedOutlineFill(properties?.GetFirstChild<Outline>());

        private static bool HasUnsupportedOutlineFill(Outline? outline) =>
            outline != null &&
            (outline.Width?.Value == 0 && outline.GetFirstChild<NoFill>() == null ||
             outline.CapType?.Value is LineCapValues cap &&
             cap != LineCapValues.Flat ||
             outline.Alignment?.Value is PenAlignmentValues alignment &&
             alignment != PenAlignmentValues.Center ||
             outline.CompoundLineType?.Value is CompoundLineValues compoundLine &&
             compoundLine != CompoundLineValues.Single ||
             outline.ChildElements.Any(child =>
                 child is not SolidFill and not NoFill));

        private static bool HasUnresolvedSeriesColor(
            C.ChartShapeProperties? properties, ColorScheme? colorScheme) {
            if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(properties?.GetFirstChild<SolidFill>()) ||
                !OfficeOpenXmlThemeColorResolver.ResolveColor(
                    properties?.GetFirstChild<SolidFill>(), colorScheme).HasValue) {
                return true;
            }

            Outline? outline = properties?.GetFirstChild<Outline>();
            if (outline == null ||
                (outline.GetFirstChild<NoFill>() == null &&
                 outline.GetFirstChild<SolidFill>() == null)) {
                return true;
            }

            return HasUnresolvedSolidFill(
                outline.GetFirstChild<SolidFill>(), colorScheme);
        }

        private static bool HasUnresolvedPointColor(
            C.ChartShapeProperties? properties, ColorScheme? colorScheme) {
            if (HasUnresolvedSolidFill(properties?.GetFirstChild<SolidFill>(), colorScheme) ||
                HasUnresolvedSolidFill(properties?.GetFirstChild<Outline>()?.GetFirstChild<SolidFill>(), colorScheme)) return true;
            if (properties?.GetFirstChild<PatternFill>() is PatternFill pattern)
                return OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(pattern.GetFirstChild<ForegroundColor>()) ||
                    OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(pattern.GetFirstChild<BackgroundColor>()) ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.GetFirstChild<ForegroundColor>(), colorScheme).HasValue ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.GetFirstChild<BackgroundColor>(), colorScheme).HasValue;
            return false;
        }

        private static bool HasUnresolvedSolidFill(
            SolidFill? fill, ColorScheme? colorScheme) =>
            fill != null &&
            (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(fill) ||
             !OfficeOpenXmlThemeColorResolver.ResolveColor(fill, colorScheme).HasValue);

        private static IReadOnlyList<OfficeColor?>? ReadBubblePointColors(
            C.BubbleChartSeries series, int pointCount, ColorScheme? colorScheme, int maximumPoints) {
            var colors = new OfficeColor?[pointCount];
            bool found = false;
            foreach (C.DataPoint point in OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(series.Elements<C.DataPoint>(), maximumPoints)) {
                uint? sourceIndex = point.GetFirstChild<C.Index>()?.Val?.Value;
                if (!sourceIndex.HasValue || sourceIndex.Value >= (uint)pointCount) continue;
                C.ChartShapeProperties? properties =
                    point.GetFirstChild<C.ChartShapeProperties>();
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                    properties?.GetFirstChild<SolidFill>(), colorScheme);
                if (!color.HasValue) continue;
                colors[(int)sourceIndex.Value] = color;
                found = true;
            }
            return found ? colors : null;
        }

        private static bool TryReadStrictCachedNumbers(OpenXmlElement? container,
            bool allowNegative, int maximumPoints, out IReadOnlyList<double> values) {
            values = Array.Empty<double>();
            if (container == null) {
                return false;
            }

            List<C.NumericPoint> points =
                OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(container.Descendants<C.NumericPoint>(), maximumPoints);
            if (points.Count == 0) {
                return false;
            }

            int length = OfficeOpenXmlChartCacheReader.GetCachedPointLength(container, points,
                point => point.Index?.Value, maximumPoints);
            if (length != points.Count) {
                return false;
            }

            var parsed = new double[length];
            var seen = new bool[length];
            for (int pointIndex = 0; pointIndex < points.Count; pointIndex++) {
                C.NumericPoint point = points[pointIndex];
                uint rawIndex = point.Index?.Value ?? (uint)pointIndex;
                if (rawIndex >= (uint)length || seen[(int)rawIndex]) {
                    return false;
                }

                string? text = point.NumericValue?.Text;
                if (!double.TryParse(text, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out double value) ||
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

    }
}
