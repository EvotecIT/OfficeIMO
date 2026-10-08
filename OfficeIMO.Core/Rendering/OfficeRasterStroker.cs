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
        var style = new OfficeStrokeOutlineOptions { StartCap = OfficeStrokeOutlineOptions.FromCap(cap), EndCap = OfficeStrokeOutlineOptions.FromCap(cap) };
        AppendOutline(points, width, join, miterLimit, pixelsPerUnit, output, style);
    }

    internal static void AppendOutline(IReadOnlyList<OfficePoint> points, double width, OfficeStrokeLineJoin join,
        double miterLimit, double pixelsPerUnit, List<List<OfficePoint>> output, OfficeStrokeOutlineOptions style) {
        if (points == null || !(width > 0) || double.IsInfinity(width)) return;
        var path = Distinct(points, style.PreserveExactGeometry, out var degenerate);
        if (path.Count == 0) return;
        var half = width / 2.0;
        var closed = path.Count >= 3 && (style.Closed ?? Same(path[0], path[path.Count - 1]));
        if (closed) {
            degenerate[0] |= degenerate[degenerate.Count - 1];
            path.RemoveAt(path.Count - 1); degenerate.RemoveAt(degenerate.Count - 1);
        }
        if (path.Count == 1) {
            AddCap(output, path[0], new OfficePoint(-1, 0), half, style.StartCap, pixelsPerUnit, style);
            AddCap(output, path[0], new OfficePoint(1, 0), half, style.EndCap, pixelsPerUnit, style);
            return;
        }

        if (double.IsNaN(miterLimit) || double.IsInfinity(miterLimit) || miterLimit < 1) miterLimit = 4;
        var count = path.Count;
        var segments = closed ? count : count - 1;
        var directions = new OfficePoint[segments];
        for (var i = 0; i < segments; i++) {
            style.CancellationToken.ThrowIfCancellationRequested();
            var a = path[i];
            var b = path[(i + 1) % count];
            var length = Math.Sqrt((b.X - a.X) * (b.X - a.X) + (b.Y - a.Y) * (b.Y - a.Y));
            var direction = new OfficePoint((b.X - a.X) / length, (b.Y - a.Y) / length);
            directions[i] = direction;
            var nx = -direction.Y * half;
            var ny = direction.X * half;
            AddPiece(output, new List<OfficePoint> { new OfficePoint(a.X + nx, a.Y + ny), new OfficePoint(b.X + nx, b.Y + ny), new OfficePoint(b.X - nx, b.Y - ny), new OfficePoint(a.X - nx, a.Y - ny) }, style);
        }

        if (closed) {
            for (var i = 0; i < count; i++) AddJoin(output, path[i], directions[(i - 1 + segments) % segments], directions[i], half, join, style.DegenerateMiterLimitOne && degenerate[i] ? 1D : miterLimit, pixelsPerUnit, style);
            return;
        }

        for (var i = 1; i < count - 1; i++) AddJoin(output, path[i], directions[i - 1], directions[i], half, join, style.DegenerateMiterLimitOne && degenerate[i] ? 1D : miterLimit, pixelsPerUnit, style);
        AddCap(output, path[0], new OfficePoint(-directions[0].X, -directions[0].Y), half, style.StartCap, pixelsPerUnit, style);
        AddCap(output, path[count - 1], directions[segments - 1], half, style.EndCap, pixelsPerUnit, style);
    }

    private static void AddJoin(List<List<OfficePoint>> output, OfficePoint vertex, OfficePoint incoming, OfficePoint outgoing, double half, OfficeStrokeLineJoin join, double miterLimit, double pixelsPerUnit, OfficeStrokeOutlineOptions style) {
        style.CancellationToken.ThrowIfCancellationRequested();
        var cross = incoming.X * outgoing.Y - incoming.Y * outgoing.X;
        var dot = incoming.X * outgoing.X + incoming.Y * outgoing.Y;
        if (Math.Abs(cross) <= Epsilon) {
            // Straight on needs nothing; a full reversal only has an outside when the join is round.
            if (dot < 0 && join == OfficeStrokeLineJoin.Round) AddCap(output, vertex, incoming, half, OfficeStrokeOutlineCap.Round, pixelsPerUnit, style);

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
            if (cosHalf > Epsilon) {
                var mx = inX + outX;
                var my = inY + outY;
                var length = Math.Sqrt(mx * mx + my * my);
                var tip = new OfficePoint(vertex.X + mx / length * half / cosHalf, vertex.Y + my / length * half / cosHalf);
                if (1 / cosHalf <= miterLimit) piece.Add(tip);
                else if (style.ClipMiter) {
                    double fraction = (miterLimit * half - half * cosHalf) / (half / cosHalf - half * cosHalf);
                    var inner = piece[1];
                    var outer = new OfficePoint(vertex.X + outX * half, vertex.Y + outY * half);
                    piece.Add(new OfficePoint(inner.X + fraction * (tip.X - inner.X), inner.Y + fraction * (tip.Y - inner.Y)));
                    piece.Add(new OfficePoint(outer.X + fraction * (tip.X - outer.X), outer.Y + fraction * (tip.Y - outer.Y)));
                }
            }
        }

        piece.Add(new OfficePoint(vertex.X + outX * half, vertex.Y + outY * half));
        AddPiece(output, piece, style);
    }

    /// <summary>Adds the cap at an end point; <paramref name="outward"/> is the unit direction pointing away from the line.</summary>
    internal static void AddCap(List<List<OfficePoint>> output, OfficePoint end, OfficePoint outward, double half, OfficeStrokeOutlineCap cap, double pixelsPerUnit, OfficeStrokeOutlineOptions style) {
        if (cap == OfficeStrokeOutlineCap.Flat) return;
        var nx = -outward.Y * half;
        var ny = outward.X * half;
        if (cap == OfficeStrokeOutlineCap.Square) {
            var ox = outward.X * half;
            var oy = outward.Y * half;
            AddPiece(output, new List<OfficePoint> { new OfficePoint(end.X + nx, end.Y + ny), new OfficePoint(end.X + nx + ox, end.Y + ny + oy), new OfficePoint(end.X - nx + ox, end.Y - ny + oy), new OfficePoint(end.X - nx, end.Y - ny) }, style);
            return;
        }

        if (cap == OfficeStrokeOutlineCap.Triangle) {
            AddPiece(output, new List<OfficePoint> { new OfficePoint(end.X + nx, end.Y + ny),
                new OfficePoint(end.X + outward.X * half, end.Y + outward.Y * half),
                new OfficePoint(end.X - nx, end.Y - ny) }, style);
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

        AddPiece(output, piece, style);
    }

    private static void AddPiece(List<List<OfficePoint>> output, List<OfficePoint> piece, OfficeStrokeOutlineOptions style) {
        if (piece.Count < 3) return;
        var area = 0.0;
        var origin = piece[0];
        for (int i = 0, j = piece.Count - 1; i < piece.Count; j = i++)
            area += (piece[j].X - origin.X) * (piece[i].Y - origin.Y) - (piece[i].X - origin.X) * (piece[j].Y - origin.Y);
        if (double.IsNaN(area) || (style.PreserveExactGeometry ? area == 0D : Math.Abs(area) <= Epsilon * Epsilon)) return;
        if (area < 0) piece.Reverse();
        style.ChargePoints?.Invoke(piece.Count);
        output.Add(piece);
    }

    private static List<OfficePoint> Distinct(IReadOnlyList<OfficePoint> points, bool exact, out List<bool> degenerate) {
        degenerate = new List<bool>();
        var result = new List<OfficePoint>(points.Count);
        foreach (var point in points) {
            if (double.IsNaN(point.X) || double.IsNaN(point.Y) || double.IsInfinity(point.X) || double.IsInfinity(point.Y)) continue;
            if (result.Count == 0 || !(exact ? result[result.Count - 1].Equals(point) : Same(result[result.Count - 1], point))) { result.Add(point); degenerate.Add(false); }
            else degenerate[degenerate.Count - 1] = true;
        }

        // Keep the closing point of a ring even though it equals the first: it marks the ring as closed.
        return result;
    }

    private static bool Same(OfficePoint a, OfficePoint b) => Distance(a, b) <= 0.000001;

    private static double Distance(OfficePoint a, OfficePoint b) => Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
}
