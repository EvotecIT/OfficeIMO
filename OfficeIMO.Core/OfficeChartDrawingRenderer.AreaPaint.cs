using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static void AddAreaRun(OfficeDrawing drawing, IReadOnlyList<OfficePoint> topPoints, IReadOnlyList<OfficePoint> bottomPoints, OfficeColor color, double strokeWidth, OfficeStrokeDashStyle dashStyle, bool showStroke, OfficeChartPointStyle? paint,
        ChartPlotBounds? plotBounds = null, ValueRange? valueRange = null) {
        if (topPoints.Count < 2 || bottomPoints.Count != topPoints.Count) {
            return;
        }

        var areaPoints = new List<OfficePoint>(topPoints.Count + bottomPoints.Count);
        areaPoints.AddRange(topPoints);
        for (int i = bottomPoints.Count - 1; i >= 0; i--) {
            areaPoints.Add(bottomPoints[i]);
        }

        bool clipped = plotBounds.HasValue && valueRange.HasValue;
        ChartPlotBounds bounds = plotBounds.GetValueOrDefault();
        ValueRange range = valueRange.GetValueOrDefault();
        List<OfficePoint> boundary = areaPoints;
        if (clipped)
            areaPoints = ClipValuePolygon(areaPoints, range, bounds);
        OfficeColor? outline = paint?.ShowOutline == false || !showStroke ? null : paint?.OutlineColor ?? color;
        double width = paint?.OutlineWidth?.ToPoints() ?? strokeWidth;
        if (areaPoints.Count >= 3) {
            // Paint only the original boundary below: clipping an area must not invent an outline on the clip edge.
            var fillPaint = new OfficeChartPointStyle(paint?.FillColor, paint?.NoFill ?? false,
                paint?.Hatch, paint?.HatchColor, showOutline: false);
            AddStyledPointPolygon(drawing, areaPoints, color, fillPaint, null, 0);
        }
        if (outline.HasValue) {
            if (paint?.OutlineJoin is OfficeStrokeLineJoin join) {
                AddJoinedSeriesOutline(drawing, boundary, outline.Value, width, dashStyle, join,
                    closed: true, plotBounds: clipped ? bounds : null, valueRange: clipped ? range : null);
            } else if (clipped) {
                AddClippedValueLine(drawing, topPoints, outline.Value, width, dashStyle, range, bounds);
                AddClippedValueLine(drawing, bottomPoints, outline.Value, paint?.OutlineWidth?.ToPoints() ?? 0.5D, OfficeStrokeDashStyle.Solid,
                    range, bounds);
                AddClippedValueLine(drawing, new[] { bottomPoints[0], topPoints[0] }, outline.Value, paint?.OutlineWidth?.ToPoints() ?? 0.5D,
                    OfficeStrokeDashStyle.Solid, range, bounds);
                AddClippedValueLine(drawing, new[] { topPoints[topPoints.Count - 1], bottomPoints[bottomPoints.Count - 1] },
                    outline.Value, paint?.OutlineWidth?.ToPoints() ?? 0.5D, OfficeStrokeDashStyle.Solid, range, bounds);
            } else {
                if (paint != null) AddPolygonShape(drawing, areaPoints, null, outline, width, dashStyle: dashStyle,
                    strokeLineJoin: paint.OutlineJoin);
                else {
                    AddPolygonShape(drawing, areaPoints, null, outline, 0.5D);
                    AddPointLine(drawing, topPoints, outline.Value, width, dashStyle);
                }
            }
        }
    }

}
