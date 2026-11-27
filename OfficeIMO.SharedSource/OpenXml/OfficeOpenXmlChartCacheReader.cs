using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Bounded indexed-cache decoding shared by native chart adapters.</summary>
    internal static class OfficeOpenXmlChartCacheReader {
        internal static IReadOnlyList<string> ReadCachedStrings(OfficeXmlNode container, int maximumPoints) {
            if (container.IsNone) {
                return Array.Empty<string>();
            }

            List<OfficeXmlNode> stringPoints = SortByIndex(GetBoundedCachedPoints(EnumeratePoints(container, numeric: false), maximumPoints));
            if (stringPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    stringPoints,
                    PointIndex,
                    PointText,
                    string.Empty, maximumPoints);
            }

            List<OfficeXmlNode> numericPoints = SortByIndex(GetBoundedCachedPoints(EnumeratePoints(container, numeric: true), maximumPoints));
            if (numericPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    numericPoints,
                    PointIndex,
                    PointText,
                    string.Empty, maximumPoints);
            }

            return Array.Empty<string>();
        }

        /// <summary>Reads a numeric cache. A point Office left out (an empty cell) becomes <paramref name="missingValue"/>.</summary>
        internal static IReadOnlyList<double> ReadCachedNumbers(OfficeXmlNode container, int maximumPoints, double missingValue = 0D) {
            if (container.IsNone) {
                return Array.Empty<double>();
            }

            List<OfficeXmlNode> points = SortByIndex(GetBoundedCachedPoints(EnumeratePoints(container, numeric: true), maximumPoints));
            if (points.Count == 0) {
                // A series whose cells are all empty keeps its declared length as missing values,
                // when the caller asks for missing values rather than zeros.
                uint? declared = container.FirstDescendant(NsC, "ptCount").UIntVal();
                if (!double.IsNaN(missingValue) || !declared.HasValue || declared.Value == 0U) return Array.Empty<double>();
                if (declared.Value > maximumPoints)
                    throw new InvalidDataException($"The chart cache declares more than the supported limit of {maximumPoints} points.");
                return Enumerable.Repeat(double.NaN, (int)declared.Value).ToArray();
            }

            return CreateIndexedCache(
                container,
                points,
                PointIndex,
                point => {
                OfficeXmlNode numericValue = point.Child(NsC, "v");
                string? text = numericValue.IsNone ? null : numericValue.InnerText;
                if (OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                    !double.IsNaN(value) &&
                    !double.IsInfinity(value)) {
                    return value;
                }

                return 0D;
                },
                missingValue, maximumPoints);
        }

        // The c:pt elements of string data (c:strCache, c:strLit, c:lvl) or of numeric data (c:numCache, c:numLit) below
        // the container, in document order and outside extension markup below it, read lazily so a bounded read stops
        // early.
        private static IEnumerable<OfficeXmlNode> EnumeratePoints(OfficeXmlNode container, bool numeric) {
            for (OfficeXmlNode child = container.FirstChild; !child.IsNone; child = child.NextContentInSubtree(container)) {
                if (child.Is(NsC, "pt") && IsPointOf(child.Parent, numeric)) yield return child;
            }
        }

        // c:pt is a numeric point in c:numCache and c:numLit, and a string point everywhere else it occurs (c:strCache,
        // c:strLit, c:lvl and extension caches such as c15:dlblRangeCache).
        private static bool IsPointOf(OfficeXmlNode data, bool numeric) =>
            numeric == (data.NamespaceId == NsC && data.LocalName is "numCache" or "numLit");

        // Sorts by c:idx (0 when missing) with the same comparisons as sorting the points directly, reading each index once.
        private static List<OfficeXmlNode> SortByIndex(List<OfficeXmlNode> points) {
            var keyed = new List<KeyValuePair<uint, OfficeXmlNode>>(points.Count);
            foreach (OfficeXmlNode point in points) keyed.Add(new KeyValuePair<uint, OfficeXmlNode>(PointIndex(point) ?? 0U, point));
            keyed.Sort((left, right) => left.Key.CompareTo(right.Key));
            for (int index = 0; index < keyed.Count; index++) points[index] = keyed[index].Value;
            return points;
        }

        private static uint? PointIndex(OfficeXmlNode point) => point.UIntAttr("idx");

        private static string PointText(OfficeXmlNode point) {
            OfficeXmlNode value = point.Child(NsC, "v");
            return value.IsNone ? string.Empty : value.InnerText;
        }

        internal static IReadOnlyList<TValue> CreateIndexedCache<TPoint, TValue>(
            OfficeXmlNode container,
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

        internal static int GetCachedPointLength<TPoint>(OfficeXmlNode container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex, int maximumPoints) {
            if (points.Count > maximumPoints) {
                throw new InvalidDataException($"The chart cache exceeds the supported limit of {maximumPoints} points.");
            }

            uint? pointCount = container.FirstDescendant(NsC, "ptCount").UIntVal();
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

        // SDK callers
        internal static IReadOnlyList<string> ReadCachedStrings(OpenXmlElement? container, int maximumPoints) => ReadCachedStrings(OfficeXmlNode.FromSdk(container), maximumPoints);
        internal static IReadOnlyList<double> ReadCachedNumbers(OpenXmlElement? container, int maximumPoints, double missingValue = 0D) => ReadCachedNumbers(OfficeXmlNode.FromSdk(container), maximumPoints, missingValue);
        internal static IReadOnlyList<TValue> CreateIndexedCache<TPoint, TValue>(OpenXmlElement container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex, Func<TPoint, TValue> getValue, TValue defaultValue, int maximumPoints) =>
            CreateIndexedCache(OfficeXmlNode.FromSdk(container), points, getIndex, getValue, defaultValue, maximumPoints);
        internal static int GetCachedPointLength<TPoint>(OpenXmlElement container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex, int maximumPoints) =>
            GetCachedPointLength(OfficeXmlNode.FromSdk(container), points, getIndex, maximumPoints);
    }
}
