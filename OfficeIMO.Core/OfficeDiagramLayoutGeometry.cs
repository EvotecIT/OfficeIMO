using System;

namespace OfficeIMO.Drawing;

internal readonly struct OfficeDiagramNodeBounds {
    internal OfficeDiagramNodeBounds(double x, double y, double width,
        double height) {
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    internal double X { get; }
    internal double Y { get; }
    internal double Width { get; }
    internal double Height { get; }
}

internal static class OfficeDiagramLayoutGeometry {
    internal static OfficeDiagramNodeBounds GetPyramidNodeBounds(
        int nodeCount, int nodeIndex, double width, double height) {
        if (nodeCount < 1) throw new ArgumentOutOfRangeException(nameof(nodeCount));
        if (nodeIndex < 0 || nodeIndex >= nodeCount) {
            throw new ArgumentOutOfRangeException(nameof(nodeIndex));
        }
        if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(width));
        }
        if (double.IsNaN(height) || double.IsInfinity(height) || height <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(height));
        }

        double progress = nodeCount == 1
            ? 1D
            : nodeIndex / (double)(nodeCount - 1);
        double nodeWidth = width * (0.28D + 0.5D * progress);
        double cellHeight = height * 0.82D / nodeCount;
        double nodeHeight = Math.Min(height * 0.19D, cellHeight * 0.86D);
        double centerY = height * 0.09D
            + (nodeIndex + 0.5D) * cellHeight;
        return new OfficeDiagramNodeBounds(
            (width - nodeWidth) / 2D,
            centerY - nodeHeight / 2D,
            nodeWidth,
            nodeHeight);
    }
}
