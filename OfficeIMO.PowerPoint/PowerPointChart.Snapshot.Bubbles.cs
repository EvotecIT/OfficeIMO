using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        private static PowerPointChartData? ReadBubbleSeriesData(
            IEnumerable<C.BubbleChartSeries> seriesElements, ColorScheme? colorScheme = null) {
            List<C.BubbleChartSeries> source = seriesElements.ToList();
            if (source.Count == 0) return null;

            var series = new List<PowerPointChartSeries>();
            IReadOnlyList<double>? categoryXValues = null;
            for (int seriesIndex = 0; seriesIndex < source.Count; seriesIndex++) {
                C.BubbleChartSeries element = source[seriesIndex];
                if (element.Elements<C.Trendline>().Any() ||
                    element.Elements<C.ErrorBars>().Any() ||
                    HasUnsupportedSeriesStyle(
                        element.GetFirstChild<C.ChartShapeProperties>()) ||
                    HasUnresolvedSeriesColor(
                        element.GetFirstChild<C.ChartShapeProperties>(), colorScheme) ||
                    element.Elements<C.DataPoint>().Any(point =>
                        HasUnsupportedPointStyle(
                            point.GetFirstChild<C.ChartShapeProperties>()) ||
                        HasUnresolvedPointColor(
                            point.GetFirstChild<C.ChartShapeProperties>(), colorScheme))) {
                    return null;
                }
                if (!TryReadStrictCachedNumbers(element.GetFirstChild<C.XValues>(),
                        allowNegative: true, out IReadOnlyList<double> xValues) ||
                    !TryReadStrictCachedNumbers(element.GetFirstChild<C.YValues>(),
                        allowNegative: true, out IReadOnlyList<double> yValues) ||
                    !TryReadStrictCachedNumbers(element.GetFirstChild<C.BubbleSize>(),
                        allowNegative: false, out IReadOnlyList<double> bubbleSizes)) {
                    return null;
                }
                if (xValues.Count != yValues.Count || xValues.Count != bubbleSizes.Count) {
                    return null;
                }
                C.InvertIfNegative? invertIfNegative =
                    element.GetFirstChild<C.InvertIfNegative>();
                if (yValues.Any(value => value < 0D) &&
                    invertIfNegative != null &&
                    invertIfNegative.Val?.Value != false) {
                    return null;
                }

                int pointCount = xValues.Count;
                if (pointCount == 0) continue;

                IReadOnlyList<double> normalizedY = NormalizeValues(yValues, pointCount);
                IReadOnlyList<double> normalizedSizes = NormalizeValues(bubbleSizes, pointCount);
                if (normalizedY.Count == 0 || normalizedSizes.Count == 0) continue;

                IReadOnlyList<double> normalizedX = xValues.Take(pointCount).ToList();
                categoryXValues ??= normalizedX;
                string name = ReadSeriesName(element);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (seriesIndex + 1).ToString(CultureInfo.InvariantCulture);
                }

                var item = new PowerPointChartSeries(name, normalizedY, normalizedX,
                    PowerPointChartSnapshotKind.Bubble,
                    ReadSeriesColor(element, PowerPointChartSnapshotKind.Bubble, colorScheme),
                    ReadSeriesStrokeWidth(element)) {
                    BubbleSizes = normalizedSizes,
                    PointColors = ReadBubblePointColors(element, pointCount, colorScheme),
                    StrokeColor = ReadSeriesStrokeColor(element, colorScheme),
                    ShowStroke = IsSeriesStrokeVisible(element),
                    SourceIndex = element.GetFirstChild<C.Index>()?.Val?.Value
                };
                series.Add(item);
            }

            if (series.Count == 0 || categoryXValues == null || categoryXValues.Count == 0) {
                return null;
            }

            IReadOnlyList<string> categories = categoryXValues.Select(value =>
                value.ToString(CultureInfo.InvariantCulture)).ToList();
            return new PowerPointChartData(categories, series);
        }

        private static bool HasUnsupportedFill(C.ChartShapeProperties? properties) =>
            properties?.ChildElements.Any(child =>
                child is NoFill or GradientFill or PatternFill or BlipFill or GroupFill) == true;

        private static bool HasUnsupportedEffects(C.ChartShapeProperties? properties) =>
            properties?.GetFirstChild<EffectList>()?.ChildElements.Count > 0 ||
            properties?.GetFirstChild<EffectDag>()?.ChildElements.Count > 0;

        private static bool HasUnsupportedSeriesStyle(
            C.ChartShapeProperties? properties) =>
            HasUnsupportedFill(properties) ||
            HasUnsupportedEffects(properties) ||
            HasUnsupportedOutlineFill(properties?.GetFirstChild<Outline>());

        private static bool HasUnsupportedPointStyle(
            C.ChartShapeProperties? properties) =>
            HasUnsupportedFill(properties) ||
            HasUnsupportedEffects(properties) ||
            properties?.GetFirstChild<Outline>() != null;

        private static bool HasUnsupportedOutlineFill(Outline? outline) =>
            outline != null &&
            (outline.CompoundLineType?.Value is CompoundLineValues compoundLine &&
             compoundLine != CompoundLineValues.Single ||
             outline.ChildElements.Any(child =>
                 child is GradientFill or PatternFill or BlipFill or GroupFill
                     or PresetDash or CustomDash));

        private static bool HasUnresolvedSeriesColor(
            C.ChartShapeProperties? properties, ColorScheme? colorScheme) =>
            HasUnresolvedSolidFill(properties?.GetFirstChild<SolidFill>(), colorScheme) ||
            HasUnresolvedSolidFill(
                properties?.GetFirstChild<Outline>()?.GetFirstChild<SolidFill>(),
                colorScheme);

        private static bool HasUnresolvedPointColor(
            C.ChartShapeProperties? properties, ColorScheme? colorScheme) =>
            HasUnresolvedSolidFill(properties?.GetFirstChild<SolidFill>(), colorScheme);

        private static bool HasUnresolvedSolidFill(
            SolidFill? fill, ColorScheme? colorScheme) =>
            fill != null &&
            !OfficeOpenXmlThemeColorResolver.ResolveColor(fill, colorScheme).HasValue;

        private static IReadOnlyList<OfficeColor?>? ReadBubblePointColors(
            C.BubbleChartSeries series, int pointCount, ColorScheme? colorScheme) {
            var colors = new OfficeColor?[pointCount];
            bool found = false;
            foreach (C.DataPoint point in series.Elements<C.DataPoint>()) {
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
            bool allowNegative, out IReadOnlyList<double> values) {
            values = Array.Empty<double>();
            if (container == null) {
                return false;
            }

            List<C.NumericPoint> points =
                GetBoundedCachedPoints(container.Descendants<C.NumericPoint>());
            if (points.Count == 0) {
                return false;
            }

            int length = GetCachedPointLength(container, points,
                point => point.Index?.Value);
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
