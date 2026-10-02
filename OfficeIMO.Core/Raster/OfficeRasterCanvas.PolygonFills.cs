using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeColor color) {
        double minX = points[0].X;
        double maxX = points[0].X;
        double minY = points[0].Y;
        double maxY = points[0].Y;
        for (int i = 1; i < points.Count; i++) {
            minX = Math.Min(minX, points[i].X);
            maxX = Math.Max(maxX, points[i].X);
            minY = Math.Min(minY, points[i].Y);
            maxY = Math.Max(maxY, points[i].Y);
        }

        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            _cancellationToken.ThrowIfCancellationRequested();
            for (int px = left; px <= right; px++) {
                double coverage = PolygonCoverage(points, px, py);
                if (coverage > 0D) {
                    BlendPixel(px, py, ApplyCoverage(color, coverage));
                }
            }
        }
    }

    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeLinearGradient gradient) {
        double minX = points[0].X;
        double maxX = points[0].X;
        double minY = points[0].Y;
        double maxY = points[0].Y;
        for (int i = 1; i < points.Count; i++) {
            minX = Math.Min(minX, points[i].X);
            maxX = Math.Max(maxX, points[i].X);
            minY = Math.Min(minY, points[i].Y);
            maxY = Math.Max(maxY, points[i].Y);
        }

        double width = Math.Max(0.0001D, maxX - minX);
        double height = Math.Max(0.0001D, maxY - minY);
        OfficeGradientStop start = gradient.Stops[0];
        double dx = gradient.EndX - gradient.StartX;
        double dy = gradient.EndY - gradient.StartY;
        double lengthSquared = (dx * dx) + (dy * dy);
        if (lengthSquared <= double.Epsilon) {
            FillPolygonCore(points, start.Color);
            return;
        }

        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            _cancellationToken.ThrowIfCancellationRequested();
            double ny = ((py + 0.5D) - minY) / height;
            for (int px = left; px <= right; px++) {
                double coverage = PolygonCoverage(points, px, py);
                if (coverage <= 0D) {
                    continue;
                }

                double nx = ((px + 0.5D) - minX) / width;
                double ratio = (((nx - gradient.StartX) * dx) + ((ny - gradient.StartY) * dy)) / lengthSquared;
                BlendPixel(px, py, ApplyCoverage(InterpolateGradient(gradient, Clamp(ratio, 0D, 1D)), coverage));
            }
        }
    }

    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeRadialGradient gradient) {
        double minX = points[0].X;
        double maxX = points[0].X;
        double minY = points[0].Y;
        double maxY = points[0].Y;
        for (int i = 1; i < points.Count; i++) {
            minX = Math.Min(minX, points[i].X);
            maxX = Math.Max(maxX, points[i].X);
            minY = Math.Min(minY, points[i].Y);
            maxY = Math.Max(maxY, points[i].Y);
        }

        double width = Math.Max(0.0001D, maxX - minX);
        double height = Math.Max(0.0001D, maxY - minY);
        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            _cancellationToken.ThrowIfCancellationRequested();
            double ny = ((py + 0.5D) - minY) / height;
            for (int px = left; px <= right; px++) {
                double coverage = PolygonCoverage(points, px, py);
                if (coverage <= 0D) {
                    continue;
                }

                double nx = ((px + 0.5D) - minX) / width;
                BlendPixel(px, py, ApplyCoverage(InterpolateGradient(gradient, ComputeRadialRatio(gradient, nx, ny)), coverage));
            }
        }
    }

}
