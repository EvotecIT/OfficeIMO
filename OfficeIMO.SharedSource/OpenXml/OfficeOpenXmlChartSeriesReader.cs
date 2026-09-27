using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Reads cached series into the shared model without format-specific chart types.</summary>
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal sealed class Series {
            internal Series(uint sourceIndex, OfficeChartSeries data, bool hasUnsupportedAppearance = false) {
                SourceIndex = sourceIndex; Data = data; HasUnsupportedAppearance = hasUnsupportedAppearance;
            }
            internal uint SourceIndex { get; }
            internal OfficeChartSeries Data { get; }
            internal bool HasUnsupportedAppearance { get; }
        }

        internal sealed class Result {
            internal Result(IReadOnlyList<string> categories, IReadOnlyList<Series> series) {
                Categories = categories; Series = series;
            }
            internal IReadOnlyList<string> Categories { get; }
            internal IReadOnlyList<Series> Series { get; }
            internal OfficeChartData ToData() => new OfficeChartData(Categories, Series.Select(item => item.Data));
        }

        internal static Result? ReadCategories(IEnumerable<OpenXmlCompositeElement> source,
            OfficeChartKind kind, A.ColorScheme? scheme, OfficeChartAxisGroup axisGroup, int maximumPoints,
            bool validatePlot = true, ProjectionBudget? projectionBudget = null, int maximumPointOverrides = 1_000_000) {
            List<OpenXmlCompositeElement> elements = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(source, maximumPoints);
            if (validatePlot) ValidatePlotBudget(elements.FirstOrDefault(), maximumPoints);
            projectionBudget ??= new ProjectionBudget();
            IReadOnlyList<string> categories = Array.Empty<string>();
            var populated = new List<(OpenXmlCompositeElement Element, IReadOnlyList<double> Values, int Position)>();
            int maximumLength = 0;
            for (int position = 0; position < elements.Count; position++) {
                OpenXmlCompositeElement element = elements[position];
                var values = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(element.GetFirstChild<C.Values>(), maximumPoints);
                if (values.Count == 0) continue;
                populated.Add((element, values, position));
                var labels = OfficeOpenXmlChartCacheReader.ReadCachedStrings(element.GetFirstChild<C.CategoryAxisData>(), maximumPoints);
                if (labels.Count > categories.Count) categories = labels;
                maximumLength = Math.Max(maximumLength, Math.Max(values.Count, labels.Count));
            }
            if (maximumLength == 0) return null;
            if (categories.Count < maximumLength) {
                categories = Enumerable.Range(0, maximumLength).Select(index => index < categories.Count ? categories[index] :
                    "Category " + (index + 1).ToString(CultureInfo.InvariantCulture)).ToArray();
            }
            var series = new List<Series>();
            foreach (var item in populated) {
                projectionBudget.Reserve(categories.Count);
                double[] normalized = new double[categories.Count];
                for (int index = 0; index < item.Values.Count; index++) normalized[index] = item.Values[index];
                series.Add(ReadSeries(item.Element, normalized, null, kind, scheme, axisGroup, item.Position, maximumPoints, maximumPointOverrides));
            }
            return series.Count == 0 ? null : new Result(categories, series);
        }

        internal static Result? ReadScatter(IEnumerable<C.ScatterChartSeries> source, A.ColorScheme? scheme, int maximumPoints,
            bool validatePlot = true, ProjectionBudget? projectionBudget = null, int maximumPointOverrides = 1_000_000) {
            var series = new List<Series>();
            IReadOnlyList<string>? categories = null;
            int sourcePosition = -1;
            projectionBudget ??= new ProjectionBudget();
            var elements = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(source, maximumPoints);
            if (validatePlot) ValidatePlotBudget(elements.FirstOrDefault(), maximumPoints);
            foreach (C.ScatterChartSeries element in elements) {
                sourcePosition++;
                var x = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(element.GetFirstChild<C.XValues>(), maximumPoints);
                var y = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(element.GetFirstChild<C.YValues>(), maximumPoints);
                int count = Math.Min(x.Count, y.Count);
                if (count == 0) continue;
                projectionBudget.Reserve(Math.Max(x.Count, y.Count));
                double[] alignedX = x.Take(count).ToArray();
                categories ??= alignedX.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray();
                series.Add(ReadSeries(element, y.Take(count).ToArray(), alignedX, OfficeChartKind.Scatter,
                    scheme, OfficeChartAxisGroup.Primary, sourcePosition, maximumPoints, maximumPointOverrides));
            }
            return categories == null || series.Count == 0 ? null : new Result(categories, series);
        }

        /// <summary>Bounds normalized positions retained by one complete chart projection.</summary>
        internal sealed class ProjectionBudget {
            private long _positions;
            internal void Reserve(int positions) {
                _positions += positions;
                ValidateTotalPoints(_positions);
            }
        }

        private static void ValidateTotalPoints(long total) {
            if (total > OfficeOpenXmlChartWriter.MaximumSharedChartPoints)
                throw new InvalidDataException("The chart series exceed the supported aggregate point limit.");
        }

        private static void ValidatePlotBudget(OpenXmlElement? series, int maximumPoints) {
            if (series?.Parent?.Parent is not C.PlotArea plot) return;
            ValidatePlotBudget(plot, maximumPoints);
        }

        internal static void ValidatePlotBudget(C.PlotArea plot, int maximumPoints) {
            long total = 0;
            foreach (OpenXmlCompositeElement layer in plot.ChildElements.OfType<OpenXmlCompositeElement>()) {
                foreach (OpenXmlCompositeElement item in layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(child => child.LocalName == "ser")) {
                    int length = 0;
                    foreach (OpenXmlElement cache in item.ChildElements.Where(child => child is C.Values || child is C.XValues || child is C.YValues || child is C.BubbleSize || child is C.CategoryAxisData)) {
                        var points = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(cache.Descendants<C.NumericPoint>(), maximumPoints);
                        length = Math.Max(length, OfficeOpenXmlChartCacheReader.GetCachedPointLength(cache, points, point => point.Index?.Value, maximumPoints));
                        var strings = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(cache.Descendants<C.StringPoint>(), maximumPoints);
                        length = Math.Max(length, OfficeOpenXmlChartCacheReader.GetCachedPointLength(cache, strings, point => point.Index?.Value, maximumPoints));
                    }
                    total += length;
                    ValidateTotalPoints(total);
                }
            }
        }

        private static Series ReadSeries(OpenXmlCompositeElement element, IReadOnlyList<double> values,
            IReadOnlyList<double>? xValues, OfficeChartKind kind, A.ColorScheme? scheme,
            OfficeChartAxisGroup axisGroup, int fallbackIndex, int maximumPoints, int maximumPointOverrides) {
            C.ChartShapeProperties? properties = element.GetFirstChild<C.ChartShapeProperties>();
            A.Outline? outline = properties?.GetFirstChild<A.Outline>();
            OfficeColor? fill = OfficeOpenXmlThemeColorResolver.ResolveColor(properties?.GetFirstChild<A.SolidFill>(), scheme);
            OfficeColor? stroke = OfficeOpenXmlThemeColorResolver.ResolveColor(outline?.GetFirstChild<A.SolidFill>(), scheme);
            bool filled = kind == OfficeChartKind.Pie || kind == OfficeChartKind.Doughnut || kind == OfficeChartKind.Bubble ||
                kind == OfficeChartKind.Area || kind == OfficeChartKind.AreaStacked || kind == OfficeChartKind.AreaStacked100 ||
                kind == OfficeChartKind.ColumnClustered || kind == OfficeChartKind.ColumnStacked || kind == OfficeChartKind.ColumnStacked100 ||
                kind == OfficeChartKind.BarClustered || kind == OfficeChartKind.BarStacked || kind == OfficeChartKind.BarStacked100;
            double? width = outline?.Width?.Value is int emus && emus > 0 ? emus / 12700d : null;
            if (width > OfficeChartStyleBounds.MaximumLineWidthPoints)
                throw new InvalidDataException("The native chart line width exceeds the supported DrawingML bound.");
            C.Marker? marker = element.GetFirstChild<C.Marker>();
            OfficeColor? markerFill = OfficeOpenXmlThemeColorResolver.ResolveColor(marker?.ChartShapeProperties?.GetFirstChild<A.SolidFill>(), scheme);
            int? markerSize = marker?.Size?.Val?.Value;
            if (markerSize.HasValue && (markerSize < 2 || markerSize > 72))
                throw new InvalidDataException("The native chart marker size is outside the supported DrawingML bound.");
            OfficeChartMarkerShape? markerShape = Enum.TryParse(marker?.Symbol?.Val?.InnerText,
                ignoreCase: true, out OfficeChartMarkerShape parsedShape) ? parsedShape : null;
            A.Outline? markerOutline = marker?.ChartShapeProperties?.GetFirstChild<A.Outline>();
            OfficeColor? markerOutlineColor = OfficeOpenXmlThemeColorResolver.ResolveColor(markerOutline?.GetFirstChild<A.SolidFill>(), scheme);
            double? markerOutlineWidth = markerOutline?.Width?.Value is int markerEmus && markerEmus > 0 ? markerEmus / 12700d : null;
            if (markerOutlineWidth > OfficeChartStyleBounds.MaximumLineWidthPoints)
                throw new InvalidDataException("The native chart marker width exceeds the supported DrawingML bound.");
            C.SeriesText? text = element.GetFirstChild<C.SeriesText>();
            string name = text?.GetFirstChild<C.NumericValue>()?.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name)) {
                var cache = OfficeOpenXmlChartCacheReader.ReadCachedStrings(text, maximumPoints);
                name = cache.FirstOrDefault() ?? string.Concat(text?.Descendants<A.Text>().Select(item => item.Text) ?? Enumerable.Empty<string>());
            }
            if (string.IsNullOrWhiteSpace(name)) name = "Series " + (fallbackIndex + 1).ToString(CultureInfo.InvariantCulture);
            OfficeColor?[]? pointColors = null;
            var pointOverrides = OfficeOpenXmlChartPointStyles.GetBoundedPoints(element, maximumPointOverrides);
            foreach (C.DataPoint point in pointOverrides) {
                uint? index = point.Index?.Val?.Value;
                if (!index.HasValue || index.Value >= values.Count) continue;
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(point.ChartShapeProperties?.GetFirstChild<A.SolidFill>(), scheme);
                if (!color.HasValue) continue;
                pointColors ??= new OfficeColor?[values.Count];
                pointColors[(int)index.Value] = color;
            }
            C.ScatterStyleValues? scatterStyle = (element.Parent as C.ScatterChart)?.ScatterStyle?.Val?.Value;
            bool inheritedMarkers = (element.Parent?.GetFirstChild<C.ShowMarker>()?.Val?.Value ?? true) &&
                scatterStyle != C.ScatterStyleValues.Line && scatterStyle != C.ScatterStyleValues.Smooth;
            if (element.Parent is C.RadarChart radar)
                inheritedMarkers = radar.RadarStyle?.Val?.Value == C.RadarStyleValues.Marker;
            bool showMarkers = markerShape.HasValue ? markerShape != OfficeChartMarkerShape.None : inheritedMarkers;
            bool connectLine = outline?.GetFirstChild<A.NoFill>() == null && scatterStyle != C.ScatterStyleValues.Marker;
            bool unsupported = markerShape == OfficeChartMarkerShape.Picture ||
                outline?.GetFirstChild<A.PresetDash>() != null && ReadDash(outline) == null;
            // The shared model has one series colour and straight connecting lines.
            // Reject appearance that would otherwise be silently flattened in an export.
            C.Smooth? smoothing = element.GetFirstChild<C.Smooth>() ?? element.Parent?.GetFirstChild<C.Smooth>();
            bool curved = smoothing != null ? smoothing.Val?.Value != false :
                scatterStyle == C.ScatterStyleValues.Smooth || scatterStyle == C.ScatterStyleValues.SmoothMarker;
            unsupported |= connectLine && curved;
            unsupported |= showMarkers && marker?.ChartShapeProperties?.GetFirstChild<A.NoFill>() != null;
            unsupported |= showMarkers && markerOutline?.GetFirstChild<A.NoFill>() != null;
            unsupported |= !filled && connectLine && markerFill.HasValue &&
                (!stroke.HasValue || showMarkers && stroke.Value != markerFill.Value);
            OfficeColor? seriesColor = filled ? fill : !connectLine && showMarkers ? markerFill ?? stroke ?? fill : stroke ?? fill;
            var data = new OfficeChartSeries(name, values, xValues, seriesColor,
                pointColors, showMarkers: showMarkers,
                connectLine: connectLine,
                markerSize: markerSize, markerShape: markerShape,
                markerOutlineColor: markerOutlineColor, markerOutlineWidth: markerOutlineWidth,
                strokeWidth: width, strokeDashStyle: ReadDash(outline), renderKind: kind, axisGroup: axisGroup)
                .WithPointStyles(OfficeOpenXmlChartPointStyles.Read(pointOverrides, values.Count, scheme));
            return new Series(element.GetFirstChild<C.Index>()?.Val?.Value ?? (uint)fallbackIndex, data, unsupported);
        }

        private static OfficeStrokeDashStyle? ReadDash(A.Outline? outline) => outline?.GetFirstChild<A.PresetDash>()?.Val?.InnerText switch {
            null => null,
            "solid" => OfficeStrokeDashStyle.Solid,
            "dash" => OfficeStrokeDashStyle.Dash,
            "dot" => OfficeStrokeDashStyle.Dot,
            "dashDot" => OfficeStrokeDashStyle.DashDot,
            "lgDashDotDot" => OfficeStrokeDashStyle.DashDotDot,
            _ => null
        };
    }
}
