using System;
using System.Collections.Generic;



namespace OfficeIMO.Drawing;

/// <summary>
/// Turns polylines into stroke outlines. An outline is a set of closed pieces (segment bodies,
/// joins, and caps) that all wind the same way, so their non-zero union is the stroke: one fill
/// paints it with the coverage of any other shape and no pixel is blended twice.
/// </summary>
internal static class OfficeRasterStroker {
    private const double Epsilon = 0.000000001;

    /// <summary>
    /// Appends the outline of one polyline. A polyline whose last point repeats its first is closed
    /// and gets a join there instead of caps. <paramref name="pixelsPerUnit"/> sizes round joins and caps.
    /// </summary>
    internal static void AppendOutline(IReadOnlyList<OfficePoint> points, double width, OfficeStrokeLineCap cap, OfficeStrokeLineJoin join, double miterLimit, double pixelsPerUnit, List<List<OfficePoint>> output) {
        if (points == null || !(width > 0) || double.IsInfinity(width)) return;
        var path = Distinct(points);
        if (path.Count == 0) return;
        var half = width / 2.0;
        var closed = path.Count >= 3 && Same(path[0], path[path.Count - 1]);
        if (closed) path.RemoveAt(path.Count - 1);
        if (path.Count == 1) {
            if (cap == OfficeStrokeLineCap.Round) AddPiece(output, OfficeCurveFlattening.Ellipse(path[0].X, path[0].Y, half, half, pixelsPerUnit));
            else if (cap == OfficeStrokeLineCap.Square) AddPiece(output, new List<OfficePoint> { new OfficePoint(path[0].X - half, path[0].Y - half), new OfficePoint(path[0].X + half, path[0].Y - half), new OfficePoint(path[0].X + half, path[0].Y + half), new OfficePoint(path[0].X - half, path[0].Y + half) });
            return;
        }

        if (double.IsNaN(miterLimit) || double.IsInfinity(miterLimit) || miterLimit < 1) miterLimit = 4;
        var count = path.Count;
        var segments = closed ? count : count - 1;
        var directions = new OfficePoint[segments];
        for (var i = 0; i < segments; i++) {
            var a = path[i];
            var b = path[(i + 1) % count];
            var length = Math.Sqrt((b.X - a.X) * (b.X - a.X) + (b.Y - a.Y) * (b.Y - a.Y));
            var direction = new OfficePoint((b.X - a.X) / length, (b.Y - a.Y) / length);
            directions[i] = direction;
            var nx = -direction.Y * half;
            var ny = direction.X * half;
            AddPiece(output, new List<OfficePoint> { new OfficePoint(a.X + nx, a.Y + ny), new OfficePoint(b.X + nx, b.Y + ny), new OfficePoint(b.X - nx, b.Y - ny), new OfficePoint(a.X - nx, a.Y - ny) });
        }

        if (closed) {
            for (var i = 0; i < count; i++) AddJoin(output, path[i], directions[(i - 1 + segments) % segments], directions[i], half, join, miterLimit, pixelsPerUnit);
            return;
        }

        for (var i = 1; i < count - 1; i++) AddJoin(output, path[i], directions[i - 1], directions[i], half, join, miterLimit, pixelsPerUnit);
        AddCap(output, path[0], new OfficePoint(-directions[0].X, -directions[0].Y), half, cap, pixelsPerUnit);
        AddCap(output, path[count - 1], directions[segments - 1], half, cap, pixelsPerUnit);
    }

    private static void AddJoin(List<List<OfficePoint>> output, OfficePoint vertex, OfficePoint incoming, OfficePoint outgoing, double half, OfficeStrokeLineJoin join, double miterLimit, double pixelsPerUnit) {
        var cross = incoming.X * outgoing.Y - incoming.Y * outgoing.X;
        var dot = incoming.X * outgoing.X + incoming.Y * outgoing.Y;
        if (Math.Abs(cross) <= Epsilon) {
            // Straight on needs nothing; a full reversal only has an outside when the join is round.
            if (dot < 0 && join == OfficeStrokeLineJoin.Round) AddCap(output, vertex, incoming, half, OfficeStrokeLineCap.Round, pixelsPerUnit);
            return;
        }

        var side = cross > 0 ? -1.0 : 1.0;
        var inX = -incoming.Y * side;
        var inY = incoming.X * side;
        var outX = -outgoing.Y * side;
        var outY = outgoing.X * side;
        var piece = new List<OfficePoint> { vertex, new OfficePoint(vertex.X + inX * half, vertex.Y + inY * half) };
        if (join == OfficeStrokeLineJoin.Round) {
            var start = Math.Atan2(inY, inX);
            var sweep = Math.Atan2(outY, outX) - start;
            if (sweep > Math.PI) sweep -= Math.PI * 2;
            else if (sweep < -Math.PI) sweep += Math.PI * 2;
            var steps = OfficeCurveFlattening.ArcSegments(half * pixelsPerUnit, sweep);
            for (var i = 1; i < steps; i++) {
                var angle = start + sweep * i / steps;
                piece.Add(new OfficePoint(vertex.X + Math.Cos(angle) * half, vertex.Y + Math.Sin(angle) * half));
            }
        } else if (join == OfficeStrokeLineJoin.Miter) {
            var cosHalf = Math.Sqrt(Math.Max(0, (1 + dot) / 2));
            if (cosHalf > Epsilon && 1 / cosHalf <= miterLimit) {
                var mx = inX + outX;
                var my = inY + outY;
                var length = Math.Sqrt(mx * mx + my * my);
                piece.Add(new OfficePoint(vertex.X + mx / length * half / cosHalf, vertex.Y + my / length * half / cosHalf));
            }
        }

        piece.Add(new OfficePoint(vertex.X + outX * half, vertex.Y + outY * half));
        AddPiece(output, piece);
    }

    /// <summary>Adds the cap at an end point; <paramref name="outward"/> is the unit direction pointing away from the line.</summary>
    private static void AddCap(List<List<OfficePoint>> output, OfficePoint end, OfficePoint outward, double half, OfficeStrokeLineCap cap, double pixelsPerUnit) {
        if (cap == OfficeStrokeLineCap.Butt) return;
        var nx = -outward.Y * half;
        var ny = outward.X * half;
        if (cap == OfficeStrokeLineCap.Square) {
            var ox = outward.X * half;
            var oy = outward.Y * half;
            AddPiece(output, new List<OfficePoint> { new OfficePoint(end.X + nx, end.Y + ny), new OfficePoint(end.X + nx + ox, end.Y + ny + oy), new OfficePoint(end.X - nx + ox, end.Y - ny + oy), new OfficePoint(end.X - nx, end.Y - ny) });
            return;
        }

        // The normal turned a quarter turn backwards is the outward direction, so the half disc sweeps that way.
        var start = Math.Atan2(ny, nx);
        var steps = Math.Max(2, OfficeCurveFlattening.ArcSegments(half * pixelsPerUnit, Math.PI));
        var piece = new List<OfficePoint>(steps + 1);
        for (var i = 0; i <= steps; i++) {
            var angle = start - Math.PI * i / steps;
            piece.Add(new OfficePoint(end.X + Math.Cos(angle) * half, end.Y + Math.Sin(angle) * half));
        }

        AddPiece(output, piece);
    }

    private static void AddPiece(List<List<OfficePoint>> output, List<OfficePoint> piece) {
        if (piece.Count < 3) return;
        var area = 0.0;
        for (int i = 0, j = piece.Count - 1; i < piece.Count; j = i++) area += piece[j].X * piece[i].Y - piece[i].X * piece[j].Y;
        if (double.IsNaN(area) || Math.Abs(area) <= Epsilon * Epsilon) return;
        if (area < 0) piece.Reverse();
        output.Add(piece);
    }

    private static List<OfficePoint> Distinct(IReadOnlyList<OfficePoint> points) {
        var result = new List<OfficePoint>(points.Count);
        foreach (var point in points) {
            if (double.IsNaN(point.X) || double.IsNaN(point.Y) || double.IsInfinity(point.X) || double.IsInfinity(point.Y)) continue;
            if (result.Count == 0 || !Same(result[result.Count - 1], point)) result.Add(point);
        }

        // Keep the closing point of a ring even though it equals the first: it marks the ring as closed.
        return result;
    }

    private static bool Same(OfficePoint a, OfficePoint b) => Distance(a, b) <= 0.000001;

    private static double Distance(OfficePoint a, OfficePoint b) => Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
}
