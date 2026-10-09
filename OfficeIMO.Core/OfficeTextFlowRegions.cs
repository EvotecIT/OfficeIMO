using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Shared conservative rectangular exclusion for native text-flow adapters.</summary>
internal static class OfficeTextFlowRegions {
    /// <summary>
    /// Selects the widest free horizontal interval in each vertical band, using
    /// the left interval to break ties. Adjacent equal intervals are coalesced.
    /// Inspection is charged to the caller's work budget before each operation.
    /// </summary>
    internal static IReadOnlyList<OfficeTextFlowRegion> Exclude(OfficeTextFlowRegion area,
        IReadOnlyList<OfficeTextFlowRegion> exclusions, Action accountInspection, CancellationToken cancellationToken) {
        var clipped = new List<OfficeTextFlowRegion>();
        var edges = new SortedSet<double> { area.Y, area.Bottom };
        foreach (OfficeTextFlowRegion exclusion in exclusions) {
            cancellationToken.ThrowIfCancellationRequested(); accountInspection();
            double left = Math.Max(area.X, exclusion.X), right = Math.Min(area.Right, exclusion.Right);
            double top = Math.Max(area.Y, exclusion.Y), bottom = Math.Min(area.Bottom, exclusion.Bottom);
            if (right <= left || bottom <= top) continue;
            clipped.Add(new OfficeTextFlowRegion(left, top, right - left, bottom - top));
            edges.Add(top); edges.Add(bottom);
        }
        var bands = new List<double>(edges);
        var result = new List<OfficeTextFlowRegion>();
        for (int band = 0; band + 1 < bands.Count; band++) {
            cancellationToken.ThrowIfCancellationRequested(); accountInspection();
            double top = bands[band], bottom = bands[band + 1];
            var obstacles = new List<OfficeTextFlowRegion>();
            foreach (OfficeTextFlowRegion exclusion in clipped) {
                cancellationToken.ThrowIfCancellationRequested(); accountInspection();
                if (exclusion.Y < bottom && exclusion.Bottom > top) obstacles.Add(exclusion);
            }
            obstacles.Sort((left, right) => left.X.CompareTo(right.X));
            double cursor = area.X, bestX = area.X, bestWidth = 0;
            foreach (OfficeTextFlowRegion obstacle in obstacles) {
                if (obstacle.X - cursor > bestWidth) { bestX = cursor; bestWidth = obstacle.X - cursor; }
                cursor = Math.Max(cursor, obstacle.Right);
            }
            if (area.Right - cursor > bestWidth) { bestX = cursor; bestWidth = area.Right - cursor; }
            if (bestWidth <= 0 || bottom <= top) continue;
            if (result.Count > 0) {
                OfficeTextFlowRegion before = result[result.Count - 1];
                if (before.X == bestX && before.Width == bestWidth && before.Bottom == top) {
                    result[result.Count - 1] = new OfficeTextFlowRegion(bestX, before.Y, bestWidth, bottom - before.Y);
                    continue;
                }
            }
            result.Add(new OfficeTextFlowRegion(bestX, top, bestWidth, bottom - top));
        }
        return result;
    }
}

/// <summary>A finite rectangle in an adapter's text coordinate space, including off-frame exclusions.</summary>
internal readonly struct OfficeTextFlowRegion {
    internal OfficeTextFlowRegion(double x, double y, double width, double height) {
        if (double.IsNaN(x) || double.IsInfinity(x) || double.IsNaN(y) || double.IsInfinity(y)
            || double.IsNaN(width) || double.IsInfinity(width) || double.IsNaN(height) || double.IsInfinity(height)
            || width < 0 || height < 0 || double.IsInfinity(x + width) || double.IsInfinity(y + height))
            throw new ArgumentOutOfRangeException(nameof(width), "Text-flow rectangles must be finite and nonnegative in size.");
        X = x; Y = y; Width = width; Height = height;
    }
    internal double X { get; }
    internal double Y { get; }
    internal double Width { get; }
    internal double Height { get; }
    internal double Right => X + Width;
    internal double Bottom => Y + Height;
}
