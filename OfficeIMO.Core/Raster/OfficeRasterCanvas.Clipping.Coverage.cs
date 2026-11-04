using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private readonly struct ClipContourSource {
        internal ClipContourSource(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule rule) {
            Contours = contours; Rule = rule;
        }
        internal IReadOnlyList<IReadOnlyList<OfficePoint>> Contours { get; }
        internal OfficeFillRule Rule { get; }
    }

    // Intersect geometric scanline intervals before accumulating area. Multiplying independent
    // pixel coverages loses where each clip lies inside the pixel and darkens coincident edges.
    private static void AccumulateClipIntersection(List<ContourCrossing> crossings, int startIndex, int count,
        IReadOnlyList<ClipContourSource> sources, int left, int right, double weight, double[] coverage) {
        var windings = new int[sources.Count];
        int outside = sources.Count, index = startIndex, endIndex = startIndex + count;
        double previous = crossings[index].X;
        while (index < endIndex) {
            double x = crossings[index].X;
            if (outside == 0 && x > previous) {
                double start = Math.Max(left, previous), end = Math.Min(right + 1D, x);
                if (end > start) for (int pixel = (int)Math.Floor(start); pixel < Math.Ceiling(end); pixel++)
                    coverage[pixel - left] += (Math.Min(end, pixel + 1D) - Math.Max(start, pixel)) * weight;
            }
            do {
                ContourCrossing crossing = crossings[index++];
                int group = crossing.Group;
                bool nonzero = sources[group].Rule == OfficeFillRule.NonZero;
                bool wasInside = nonzero ? windings[group] != 0 : (windings[group] & 1) != 0;
                windings[group] += nonzero ? crossing.WindingDelta : 1;
                bool isInside = nonzero ? windings[group] != 0 : (windings[group] & 1) != 0;
                if (wasInside != isInside) outside += isInside ? -1 : 1;
            } while (index < endIndex && Math.Abs(crossings[index].X - x) <= ContourCrossingTolerance);
            previous = x;
        }
    }

    private byte[] RasterizeClipMask(IReadOnlyList<ClipContourSource> sources, int left, int top, int width, int height) {
        ChargeIntermediateSurfacePixels(2L * width * height, MaximumClipRasterPixels);
        var image = new OfficeRasterImage(width, height);
        var maskCanvas = new OfficeRasterCanvas(image, font: null, fonts: null, cancellationToken: _cancellationToken);
        maskCanvas.ShareTransformedTextBudget(_transformedTextBudget);
        var shifted = new List<ClipContourSource>(sources.Count);
        foreach (ClipContourSource source in sources) {
            var contours = new List<IReadOnlyList<OfficePoint>>(source.Contours.Count);
            foreach (IReadOnlyList<OfficePoint> contour in source.Contours) {
                var points = new OfficePoint[contour.Count];
                for (int index = 0; index < points.Length; index++)
                    points[index] = new OfficePoint(contour[index].X - left, contour[index].Y - top);
                contours.Add(points);
            }
            shifted.Add(new ClipContourSource(contours, source.Rule));
        }
        ContourCoverageWorkspace workspace = maskCanvas.TakeContourCoverageWorkspace();
        try {
            maskCanvas.FillContourPaint(shifted[0].Contours, shifted[0].Rule, (_, _) => OfficeColor.White,
                null, workspace, shifted);
        } finally {
            maskCanvas.ReturnContourCoverageWorkspace(workspace);
        }
        byte[] pixels = image.GetPixels();
        var mask = new byte[width * height];
        for (int index = 0; index < mask.Length; index++) mask[index] = pixels[index * 4 + 3];
        return mask;
    }

    private sealed partial class OfficeRasterClipRegion {
        private byte[]? _mask;
        private int _maskLeft, _maskTop, _maskWidth, _maskHeight;
        private bool _maskBuilt;
        private OfficeRasterClipRectangle? _maskRectangle;

        internal double Coverage(OfficeRasterCanvas canvas, int x, int y) {
            if (_previous == null && _rectangle.HasValue) return _rectangle.Value.Coverage(x, y);
            if (!_maskBuilt) { BuildMask(canvas); _maskBuilt = true; }
            if (_maskRectangle.HasValue) return _maskRectangle.Value.Coverage(x, y);
            if (_mask == null) return 0D;
            int column = x - _maskLeft, row = y - _maskTop;
            return (uint)column < (uint)_maskWidth && (uint)row < (uint)_maskHeight
                ? _mask[row * _maskWidth + column] / 255D : 0D;
        }

        private void BuildMask(OfficeRasterCanvas canvas) {
            var sources = new List<ClipContourSource>();
            OfficeRasterClipRectangle? rectangles = null;
            bool allRectangles = true;
            double left = 0D, top = 0D, right = canvas.Width, bottom = canvas.Height;
            for (OfficeRasterClipRegion? region = this; region != null; region = region._previous) {
                canvas._cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<IReadOnlyList<OfficePoint>> contours;
                OfficeRasterClipRectangle? rectangle = region._rectangle;
                if (rectangle.HasValue) {
                    OfficeRasterClipRectangle r = rectangle.Value;
                    double l = r.HasFractionalEdges ? r.EdgeLeft : r.Left, t = r.HasFractionalEdges ? r.EdgeTop : r.Top;
                    double h = r.HasFractionalEdges ? r.EdgeBottom : r.Bottom, w = r.HasFractionalEdges ? r.EdgeRight : r.Right;
                    contours = new[] { new[] { new OfficePoint(l, t), new OfficePoint(w, t), new OfficePoint(w, h), new OfficePoint(l, h) } };
                } else {
                    if (region._contours == null || region._contours.Count == 0) return;
                    contours = region._contours;
                    if (OfficeTextInkClip.TryGetAxisAlignedRectangle(contours, out double l, out double t, out double w, out double h))
                        rectangle = OfficeRasterClipRectangle.FromFractionalBounds(l, t, w - l, h - t);
                }
                if (rectangle.HasValue) rectangles = rectangles.HasValue
                    ? OfficeRasterClipRectangle.Intersect(rectangles.Value, rectangle.Value) : rectangle;
                else allRectangles = false;
                double l0 = double.PositiveInfinity, t0 = double.PositiveInfinity;
                double r0 = double.NegativeInfinity, b0 = double.NegativeInfinity;
                foreach (IReadOnlyList<OfficePoint> contour in contours) {
                    if (contour.Count < 3) continue;
                    foreach (OfficePoint point in contour) {
                        canvas._cancellationToken.ThrowIfCancellationRequested();
                        if (!IsFinite(point.X) || !IsFinite(point.Y)) return;
                        l0 = Math.Min(l0, point.X); t0 = Math.Min(t0, point.Y);
                        r0 = Math.Max(r0, point.X); b0 = Math.Max(b0, point.Y);
                    }
                }
                left = Math.Max(left, l0); top = Math.Max(top, t0);
                right = Math.Min(right, r0); bottom = Math.Min(bottom, b0);
                if (right <= left || bottom <= top) return;
                sources.Add(new ClipContourSource(contours, region._fillRule));
            }
            if (allRectangles) { _maskRectangle = rectangles; return; }
            _maskLeft = (int)Math.Floor(left); _maskTop = (int)Math.Floor(top);
            _maskWidth = (int)Math.Ceiling(right) - _maskLeft; _maskHeight = (int)Math.Ceiling(bottom) - _maskTop;
            _mask = canvas.RasterizeClipMask(sources, _maskLeft, _maskTop, _maskWidth, _maskHeight);
        }
    }
}
