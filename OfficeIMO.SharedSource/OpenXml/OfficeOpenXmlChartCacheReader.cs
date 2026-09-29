using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Bounded indexed-cache decoding shared by native chart adapters.</summary>
    internal static class OfficeOpenXmlChartCacheReader {
        internal static IReadOnlyList<string> ReadCachedStrings(OpenXmlElement? container, int maximumPoints) {
            if (container == null) {
                return Array.Empty<string>();
            }

            List<C.StringPoint> stringPoints = GetBoundedCachedPoints(container.Descendants<C.StringPoint>(), maximumPoints);
            stringPoints.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (stringPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    stringPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty, maximumPoints);
            }

            List<C.NumericPoint> numericPoints = GetBoundedCachedPoints(container.Descendants<C.NumericPoint>(), maximumPoints);
            numericPoints.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (numericPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    numericPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty, maximumPoints);
            }

            return Array.Empty<string>();
        }

        internal static IReadOnlyList<double> ReadCachedNumbers(OpenXmlElement? container, int maximumPoints) {
            if (container == null) {
                return Array.Empty<double>();
            }

            List<C.NumericPoint> points = GetBoundedCachedPoints(container.Descendants<C.NumericPoint>(), maximumPoints);
            points.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (points.Count == 0) {
                return Array.Empty<double>();
            }

            return CreateIndexedCache(
                container,
                points,
                point => point.Index?.Value,
                point => {
                string? text = point.NumericValue?.Text;
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                    !double.IsNaN(value) &&
                    !double.IsInfinity(value)) {
                    return value;
                }

                return 0D;
                },
                0D, maximumPoints);
        }

        internal static IReadOnlyList<TValue> CreateIndexedCache<TPoint, TValue>(
            OpenXmlElement container,
            IReadOnlyList<TPoint> points,
            Func<TPoint, uint?> getIndex,
            Func<TPoint, TValue> getValue,
            TValue defaultValue, int maximumPoints) {
            int length = GetCachedPointLength(container, points, getIndex, maximumPoints);
            var values = Enumerable.Repeat(defaultValue, length).ToArray();
            for (int i = 0; i < points.Count; i++) {
                TPoint point = points[i];
                uint? rawIndex = getIndex(point);
                int index = rawIndex.HasValue && rawIndex.Value <= int.MaxValue
                    ? (int)rawIndex.Value
                    : i;
                if (index >= 0 && index < values.Length) {
                    values[index] = getValue(point);
                }
            }

            return values;
        }

        internal static List<TPoint> GetBoundedCachedPoints<TPoint>(IEnumerable<TPoint> points, int maximumPoints) {
            List<TPoint> boundedPoints = points
                .Take(maximumPoints + 1).ToList();
            if (boundedPoints.Count > maximumPoints) {
                throw new InvalidDataException($"The chart cache exceeds the supported limit of {maximumPoints} points.");
            }

            return boundedPoints;
        }

        internal static int GetCachedPointLength<TPoint>(OpenXmlElement container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex, int maximumPoints) {
            if (points.Count > maximumPoints) {
                throw new InvalidDataException($"The chart cache exceeds the supported limit of {maximumPoints} points.");
            }

            uint? pointCount = container.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value;
            if (pointCount > maximumPoints) {
                throw new InvalidDataException($"The chart cache declares more than the supported limit of {maximumPoints} points.");
            }

            uint maxIndex = 0U;
            bool hasIndexedPoint = false;
            for (int i = 0; i < points.Count; i++) {
                uint? index = getIndex(points[i]);
                if (!index.HasValue) {
                    continue;
                }

                if (index.Value >= maximumPoints) {
                    throw new InvalidDataException($"The chart cache point index exceeds the supported limit of {maximumPoints} points.");
                }

                hasIndexedPoint = true;
                if (index.Value > maxIndex) {
                    maxIndex = index.Value;
                }
            }

            uint indexedLength = hasIndexedPoint ? maxIndex + 1U : (uint)points.Count;
            uint length = Math.Max(pointCount ?? 0U, indexedLength);
            return (int)length;
        }

    }
}
