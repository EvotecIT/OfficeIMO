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
            OfficeChartKind kind, A.ColorScheme? scheme, OfficeChartAxisGroup axisGroup, int maximumPoints) {
            List<OpenXmlCompositeElement> elements = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(source, maximumPoints);
            IReadOnlyList<string> categories = Array.Empty<string>();
            foreach (OpenXmlCompositeElement element in elements) {
                var values = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(element.GetFirstChild<C.Values>(), maximumPoints);
                if (values.Count == 0) continue;
                categories = OfficeOpenXmlChartCacheReader.ReadCachedStrings(element.GetFirstChild<C.CategoryAxisData>(), maximumPoints);
                if (categories.Count == 0) categories = FallbackCategories(values.Count);
                break;
            }
            if (categories.Count == 0) return null;
            var series = new List<Series>();
            foreach (OpenXmlCompositeElement element in elements) {
                var values = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(element.GetFirstChild<C.Values>(), maximumPoints);
                if (values.Count == 0) continue;
                double[] normalized = new double[categories.Count];
                for (int index = 0; index < Math.Min(values.Count, normalized.Length); index++) normalized[index] = values[index];
                series.Add(ReadSeries(element, normalized, null, kind, scheme, axisGroup, series.Count, maximumPoints));
            }
            return series.Count == 0 ? null : new Result(categories, series);
        }

        internal static Result? ReadScatter(IEnumerable<C.ScatterChartSeries> source, A.ColorScheme? scheme, int maximumPoints) {
            var series = new List<Series>();
            IReadOnlyList<string>? categories = null;
            foreach (C.ScatterChartSeries element in OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(source, maximumPoints)) {
                var x = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(element.GetFirstChild<C.XValues>(), maximumPoints);
                var y = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(element.GetFirstChild<C.YValues>(), maximumPoints);
                int count = Math.Min(x.Count, y.Count);
                if (count == 0) continue;
                double[] alignedX = x.Take(count).ToArray();
                categories ??= alignedX.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray();
                series.Add(ReadSeries(element, y.Take(count).ToArray(), alignedX, OfficeChartKind.Scatter,
                    scheme, OfficeChartAxisGroup.Primary, series.Count, maximumPoints));
            }
            return categories == null || series.Count == 0 ? null : new Result(categories, series);
        }

        private static Series ReadSeries(OpenXmlCompositeElement element, IReadOnlyList<double> values,
            IReadOnlyList<double>? xValues, OfficeChartKind kind, A.ColorScheme? scheme,
            OfficeChartAxisGroup axisGroup, int fallbackIndex, int maximumPoints) {
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
            foreach (C.DataPoint point in OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(element.Elements<C.DataPoint>(), maximumPoints)) {
                uint? index = point.Index?.Val?.Value;
                if (!index.HasValue || index.Value >= values.Count) continue;
                A.SolidFill? pointFill = point.ChartShapeProperties?.GetFirstChild<A.SolidFill>();
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(pointFill, scheme);
                if (pointFill != null && !color.HasValue)
                    throw new NotSupportedException("The chart point fill cannot be resolved for a managed snapshot.");
                if (!color.HasValue) continue;
                pointColors ??= new OfficeColor?[values.Count];
                pointColors[(int)index.Value] = color;
            }
            var data = new OfficeChartSeries(name, values, xValues, filled ? fill : stroke ?? fill,
                pointColors, showMarkers: markerShape != OfficeChartMarkerShape.None,
                connectLine: outline?.GetFirstChild<A.NoFill>() == null,
                markerSize: markerSize, markerShape: markerShape,
                markerOutlineColor: markerOutlineColor, markerOutlineWidth: markerOutlineWidth,
                strokeWidth: width, strokeDashStyle: ReadDash(outline), renderKind: kind, axisGroup: axisGroup)
                .WithPointStyles(OfficeOpenXmlChartPointStyles.Read(element, values.Count, scheme));
            return new Series(element.GetFirstChild<C.Index>()?.Val?.Value ?? (uint)fallbackIndex, data,
                outline?.GetFirstChild<A.PresetDash>() != null && data.StrokeDashStyle == null);
        }

        private static IReadOnlyList<string> FallbackCategories(int count) => Enumerable.Range(1, count)
            .Select(index => "Category " + index.ToString(CultureInfo.InvariantCulture)).ToArray();

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
