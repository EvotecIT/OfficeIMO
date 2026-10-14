using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeGeometry {
    // Coordinate-system-neutral routes shared by the Draw and Visio adapters.
    internal static OfficePoint[] CreateOrthogonalConnectorRoute(OfficePoint start, OfficePoint end, bool horizontalFirst, double offset) {
        ValidateRoutingPoint(start); ValidateRoutingPoint(end); ValidateRoutingNumber(offset);
        double laneX = (start.X + end.X) / 2 + offset, laneY = (start.Y + end.Y) / 2 + offset;
        ValidateRoutingNumber(horizontalFirst ? laneX : laneY);
        return horizontalFirst
            ? new[] { start, new OfficePoint(laneX, start.Y), new OfficePoint(laneX, end.Y), end }
            : new[] { start, new OfficePoint(start.X, laneY), new OfficePoint(end.X, laneY), end };
    }

    internal static IEnumerable<OfficePoint[]> EnumerateOrthogonalConnectorRoutes(OfficePoint start, OfficePoint end, double step, int maxLanes) {
        ValidateRoutingPoint(start); ValidateRoutingPoint(end); ValidateRoutingNumber(step);
        if (maxLanes < 0) throw new ArgumentOutOfRangeException(nameof(maxLanes));
        bool primaryHorizontal = Math.Abs(end.X - start.X) < Math.Abs(end.Y - start.Y);
        foreach (double offset in ConnectorLaneOffsets(step, maxLanes)) {
            yield return CreateOrthogonalConnectorRoute(start, end, primaryHorizontal, offset);
            yield return CreateOrthogonalConnectorRoute(start, end, !primaryHorizontal, offset);
        }
        double[] offsets = ConnectorLaneOffsets(step, maxLanes).ToArray();
        foreach (double xOffset in offsets) foreach (double yOffset in offsets) {
            if (Math.Abs(xOffset) < 1e-9 && Math.Abs(yOffset) < 1e-9) continue;
            double laneX = (start.X + end.X) / 2 + xOffset, laneY = (start.Y + end.Y) / 2 + yOffset;
            ValidateRoutingNumber(laneX); ValidateRoutingNumber(laneY);
            yield return new[] { start, new OfficePoint(laneX, start.Y), new OfficePoint(laneX, laneY), new OfficePoint(end.X, laneY), end };
            yield return new[] { start, new OfficePoint(start.X, laneY), new OfficePoint(laneX, laneY), new OfficePoint(laneX, end.Y), end };
        }
    }

    private static IEnumerable<double> ConnectorLaneOffsets(double step, int maxLanes) {
        double spacing = step > 0 ? step : 0.15;
        yield return 0;
        for (int lane = 1; lane <= maxLanes; lane++) { yield return lane * spacing; yield return -lane * spacing; }
    }
    private static void ValidateRoutingPoint(OfficePoint point) { ValidateRoutingNumber(point.X); ValidateRoutingNumber(point.Y); }
    private static void ValidateRoutingNumber(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentException("Route coordinates and lane offsets must be finite.");
    }
}
