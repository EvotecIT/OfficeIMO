using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Renders semantic diagram snapshots as reusable drawing primitives.</summary>
public static class OfficeDiagramDrawingRenderer {
    private static readonly OfficeColor[] NodeColors = {
        OfficeColor.FromRgb(37, 99, 235),
        OfficeColor.FromRgb(14, 116, 144),
        OfficeColor.FromRgb(79, 70, 229),
        OfficeColor.FromRgb(2, 132, 199)
    };

    /// <summary>Renders a semantic diagram into a fixed-size vector drawing.</summary>
    public static OfficeDrawing Render(OfficeDiagramSnapshot snapshot) =>
        Render(snapshot, includeBackground: true);

    /// <summary>Renders a semantic diagram into a fixed-size vector drawing.</summary>
    /// <param name="snapshot">Diagram content and layout kind.</param>
    /// <param name="includeBackground">Whether to add the standalone white canvas background.</param>
    public static OfficeDrawing Render(OfficeDiagramSnapshot snapshot,
        bool includeBackground) {
        if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
        var drawing = new OfficeDrawing(snapshot.WidthPoints,
            snapshot.HeightPoints);
        if (includeBackground) {
            drawing.AddShape(new OfficeShape {
                Kind = OfficeShapeKind.Rectangle,
                Width = snapshot.WidthPoints,
                Height = snapshot.HeightPoints,
                FillColor = OfficeColor.White
            }, 0D, 0D);
        }

        List<NodeBox> nodes = snapshot.Kind switch {
            OfficeDiagramKind.Hierarchy => LayoutHierarchy(snapshot),
            OfficeDiagramKind.Cycle => LayoutCycle(snapshot),
            OfficeDiagramKind.List => LayoutList(snapshot),
            OfficeDiagramKind.Matrix => LayoutMatrix(snapshot),
            OfficeDiagramKind.Pyramid => LayoutPyramid(snapshot),
            OfficeDiagramKind.Relationship => LayoutRelationship(snapshot),
            _ => LayoutProcess(snapshot)
        };
        AddConnectors(drawing, snapshot, nodes);
        AddNodes(drawing, snapshot, nodes);
        return drawing;
    }

    private static List<NodeBox> LayoutProcess(
        OfficeDiagramSnapshot snapshot) {
        int count = snapshot.Nodes.Count;
        double ratio = snapshot.WidthPoints / snapshot.HeightPoints;
        int columns = Math.Max(1, Math.Min(count,
            checked((int)Math.Ceiling(Math.Sqrt(count * ratio)))));
        int rows = checked((count + columns - 1) / columns);
        double cellWidth = snapshot.WidthPoints / columns;
        double cellHeight = snapshot.HeightPoints / rows;
        double nodeWidth = Math.Max(12D, cellWidth * 0.7D);
        double nodeHeight = Math.Max(10D, cellHeight * 0.54D);
        nodeWidth = Math.Min(nodeWidth, Math.Max(1D, cellWidth - 8D));
        nodeHeight = Math.Min(nodeHeight, Math.Max(1D, cellHeight - 8D));
        var result = new List<NodeBox>(count);
        for (int index = 0; index < count; index++) {
            int row = index / columns;
            int column = index % columns;
            result.Add(new NodeBox(
                column * cellWidth + (cellWidth - nodeWidth) / 2D,
                row * cellHeight + (cellHeight - nodeHeight) / 2D,
                nodeWidth, nodeHeight));
        }
        return result;
    }

    private static List<NodeBox> LayoutHierarchy(
        OfficeDiagramSnapshot snapshot) {
        int count = snapshot.Nodes.Count;
        double rootBandHeight = snapshot.HeightPoints * 0.44D;
        double rootWidth = Math.Min(snapshot.WidthPoints * 0.36D,
            Math.Max(1D, snapshot.WidthPoints - 8D));
        double rootHeight = Math.Min(snapshot.HeightPoints * 0.24D,
            Math.Max(1D, rootBandHeight - 8D));
        var result = new List<NodeBox>(count) {
            new NodeBox((snapshot.WidthPoints - rootWidth) / 2D,
                snapshot.HeightPoints * 0.22D - rootHeight / 2D,
                rootWidth, rootHeight)
        };
        if (count == 1) return result;

        int childCount = count - 1;
        int columns = Math.Max(1, Math.Min(4, childCount));
        int rows = checked((childCount + columns - 1) / columns);
        double childRegionTop = snapshot.HeightPoints * 0.5D;
        double childRegionHeight = snapshot.HeightPoints * 0.45D;
        double cellWidth = snapshot.WidthPoints / columns;
        double cellHeight = childRegionHeight / rows;
        double childWidth = Math.Min(cellWidth * 0.7744D,
            Math.Max(1D, cellWidth - 8D));
        double childHeight = Math.Min(snapshot.HeightPoints * 0.24D,
            Math.Min(cellHeight * 0.82D, Math.Max(1D, cellHeight - 8D)));
        for (int childIndex = 0; childIndex < childCount; childIndex++) {
            int column = childIndex % columns;
            int row = childIndex / columns;
            result.Add(new NodeBox(
                column * cellWidth + (cellWidth - childWidth) / 2D,
                childRegionTop + row * cellHeight
                    + (cellHeight - childHeight) / 2D,
                childWidth, childHeight));
        }
        return result;
    }

    private static List<NodeBox> LayoutCycle(
        OfficeDiagramSnapshot snapshot) {
        int count = snapshot.Nodes.Count;
        double nodeWidth = Math.Min(snapshot.WidthPoints * 0.25D,
            Math.Max(16D, snapshot.WidthPoints / Math.Max(2.5D, count)));
        double nodeHeight = Math.Min(snapshot.HeightPoints * 0.22D,
            Math.Max(12D, snapshot.HeightPoints / Math.Max(3D, count)));
        nodeWidth = Math.Min(nodeWidth, snapshot.WidthPoints);
        nodeHeight = Math.Min(nodeHeight, snapshot.HeightPoints);
        double radiusX = Math.Max(0D,
            (snapshot.WidthPoints - nodeWidth) * 0.43D);
        double radiusY = Math.Max(0D,
            (snapshot.HeightPoints - nodeHeight) * 0.39D);
        double centerX = snapshot.WidthPoints / 2D;
        double centerY = snapshot.HeightPoints / 2D;
        var result = new List<NodeBox>(count);
        for (int index = 0; index < count; index++) {
            double angle = -Math.PI / 2D + 2D * Math.PI * index / count;
            result.Add(new NodeBox(
                centerX + radiusX * Math.Cos(angle) - nodeWidth / 2D,
                centerY + radiusY * Math.Sin(angle) - nodeHeight / 2D,
                nodeWidth, nodeHeight));
        }
        return result;
    }

    private static List<NodeBox> LayoutList(
        OfficeDiagramSnapshot snapshot) {
        int count = snapshot.Nodes.Count;
        double cellHeight = snapshot.HeightPoints / count;
        double nodeWidth = Math.Max(1D, snapshot.WidthPoints - 20D);
        double nodeHeight = Math.Min(Math.Max(10D, cellHeight * 0.68D),
            Math.Max(1D, cellHeight - 6D));
        var result = new List<NodeBox>(count);
        for (int index = 0; index < count; index++) {
            result.Add(new NodeBox(
                (snapshot.WidthPoints - nodeWidth) / 2D,
                index * cellHeight + (cellHeight - nodeHeight) / 2D,
                nodeWidth, nodeHeight));
        }
        return result;
    }

    private static List<NodeBox> LayoutMatrix(
        OfficeDiagramSnapshot snapshot) {
        int count = snapshot.Nodes.Count;
        int columns = Math.Max(1, checked((int)Math.Ceiling(Math.Sqrt(count))));
        int rows = checked((count + columns - 1) / columns);
        double cellWidth = snapshot.WidthPoints / columns;
        double cellHeight = snapshot.HeightPoints / rows;
        double nodeWidth = Math.Min(Math.Max(12D, cellWidth * 0.78D),
            Math.Max(1D, cellWidth - 8D));
        double nodeHeight = Math.Min(Math.Max(10D, cellHeight * 0.7D),
            Math.Max(1D, cellHeight - 8D));
        var result = new List<NodeBox>(count);
        for (int index = 0; index < count; index++) {
            int row = index / columns;
            int column = index % columns;
            result.Add(new NodeBox(
                column * cellWidth + (cellWidth - nodeWidth) / 2D,
                row * cellHeight + (cellHeight - nodeHeight) / 2D,
                nodeWidth, nodeHeight));
        }
        return result;
    }

    private static List<NodeBox> LayoutPyramid(
        OfficeDiagramSnapshot snapshot) {
        int count = snapshot.Nodes.Count;
        var result = new List<NodeBox>(count);
        for (int index = 0; index < count; index++) {
            OfficeDiagramNodeBounds bounds =
                OfficeDiagramLayoutGeometry.GetPyramidNodeBounds(
                    count, index, snapshot.WidthPoints,
                    snapshot.HeightPoints);
            result.Add(new NodeBox(bounds.X, bounds.Y,
                bounds.Width, bounds.Height));
        }
        return result;
    }

    private static List<NodeBox> LayoutRelationship(
        OfficeDiagramSnapshot snapshot) {
        int count = snapshot.Nodes.Count;
        double nodeWidth = Math.Min(snapshot.WidthPoints * 0.27D,
            Math.Max(16D, snapshot.WidthPoints / Math.Max(2.8D, count)));
        double nodeHeight = Math.Min(snapshot.HeightPoints * 0.25D,
            Math.Max(12D, snapshot.HeightPoints / Math.Max(3D, count)));
        nodeWidth = Math.Min(nodeWidth, snapshot.WidthPoints);
        nodeHeight = Math.Min(nodeHeight, snapshot.HeightPoints);
        double centerX = snapshot.WidthPoints / 2D;
        double centerY = snapshot.HeightPoints / 2D;
        var result = new List<NodeBox>(count) {
            new NodeBox(centerX - nodeWidth / 2D, centerY - nodeHeight / 2D,
                nodeWidth, nodeHeight)
        };
        if (count == 1) return result;
        double radiusX = Math.Max(0D, (snapshot.WidthPoints - nodeWidth) * 0.43D);
        double radiusY = Math.Max(0D, (snapshot.HeightPoints - nodeHeight) * 0.4D);
        for (int index = 1; index < count; index++) {
            double angle = -Math.PI / 2D + 2D * Math.PI * (index - 1) / (count - 1);
            result.Add(new NodeBox(
                centerX + radiusX * Math.Cos(angle) - nodeWidth / 2D,
                centerY + radiusY * Math.Sin(angle) - nodeHeight / 2D,
                nodeWidth, nodeHeight));
        }
        return result;
    }

    private static void AddConnectors(OfficeDrawing drawing,
        OfficeDiagramSnapshot snapshot, IReadOnlyList<NodeBox> nodes) {
        if (nodes.Count < 2) return;
        OfficeDiagramKind kind = snapshot.Kind;
        OfficeColor connectorColor = snapshot.Style?.ConnectorColor
            ?? OfficeColor.FromRgb(100, 116, 139);
        if (kind == OfficeDiagramKind.Hierarchy) {
            for (int index = 1; index < nodes.Count; index++) {
                AddConnector(drawing, nodes[0], nodes[index],
                    ellipticalNodes: false, connectorColor: connectorColor,
                    includeArrowHead: false);
            }
            return;
        }
        if (kind != OfficeDiagramKind.Cycle) return;
        int connectorCount = nodes.Count;
        for (int index = 0; index < connectorCount; index++) {
            AddConnector(drawing, nodes[index],
                nodes[(index + 1) % nodes.Count],
                true, connectorColor, includeArrowHead: true);
        }
    }

    private static void AddConnector(OfficeDrawing drawing, NodeBox source,
        NodeBox target, bool ellipticalNodes, OfficeColor connectorColor,
        bool includeArrowHead) {
        double deltaX = target.CenterX - source.CenterX;
        double deltaY = target.CenterY - source.CenterY;
        if (Math.Abs(deltaX) < 0.000001D
            && Math.Abs(deltaY) < 0.000001D) return;
        OfficePoint start = IntersectNodeBoundary(source, deltaX, deltaY,
            ellipticalNodes);
        OfficePoint end = IntersectNodeBoundary(target, -deltaX, -deltaY,
            ellipticalNodes);
        double x1 = start.X;
        double y1 = start.Y;
        double x2 = end.X;
        double y2 = end.Y;
        OfficeShape line = OfficeShape.Line(x1, y1, x2, y2);
        line.StrokeColor = connectorColor;
        line.StrokeWidth = 1.5D;
        if (includeArrowHead) {
            line.StrokeEndMarker = new OfficeLineMarker(
                OfficeLineMarkerKind.Triangle, 5D, 5D);
        }
        drawing.AddShape(line, Math.Min(x1, x2), Math.Min(y1, y2));
    }

    private static OfficePoint IntersectNodeBoundary(NodeBox node,
        double deltaX, double deltaY, bool ellipse) {
        double radiusX = node.Width / 2D;
        double radiusY = node.Height / 2D;
        double scale;
        if (ellipse) {
            scale = 1D / Math.Sqrt(
                deltaX * deltaX / (radiusX * radiusX)
                + deltaY * deltaY / (radiusY * radiusY));
        } else {
            scale = 1D / Math.Max(
                Math.Abs(deltaX) / radiusX,
                Math.Abs(deltaY) / radiusY);
        }
        return new OfficePoint(node.CenterX + deltaX * scale,
            node.CenterY + deltaY * scale);
    }

    private static void AddNodes(OfficeDrawing drawing,
        OfficeDiagramSnapshot snapshot, IReadOnlyList<NodeBox> nodes) {
        for (int index = 0; index < nodes.Count; index++) {
            NodeBox node = nodes[index];
            OfficeShape shape;
            if (snapshot.Kind == OfficeDiagramKind.Cycle
                || snapshot.Kind == OfficeDiagramKind.Relationship) {
                shape = OfficeShape.Ellipse(node.Width, node.Height);
            } else if (snapshot.Kind == OfficeDiagramKind.Pyramid) {
                double inset = Math.Min(node.Width * 0.13D, node.Height * 0.4D);
                shape = OfficeShape.Polygon(
                    new OfficePoint(inset, 0D),
                    new OfficePoint(node.Width - inset, 0D),
                    new OfficePoint(node.Width, node.Height),
                    new OfficePoint(0D, node.Height));
            } else {
                shape = OfficeShape.RoundedRectangle(node.Width, node.Height,
                    Math.Min(8D, Math.Min(node.Width, node.Height) * 0.18D));
            }
            IReadOnlyList<OfficeColor> nodeColors = snapshot.Style?.NodeColors
                ?? NodeColors;
            shape.FillColor = nodeColors[index % nodeColors.Count];
            shape.StrokeColor = snapshot.Style?.NodeOutlineColor
                ?? OfficeColor.White;
            shape.StrokeWidth = 1.25D;
            drawing.AddShape(shape, node.X, node.Y);
            double fontSize = Math.Max(6D, Math.Min(12D,
                Math.Min(node.Height * 0.28D, node.Width * 0.09D)));
            drawing.AddText(snapshot.Nodes[index], node.X + 3D,
                node.Y + 2D, Math.Max(1D, node.Width - 6D),
                Math.Max(1D, node.Height - 4D),
                new OfficeFontInfo(snapshot.Style?.FontFamily ?? "Calibri",
                    fontSize, OfficeFontStyle.Bold),
                snapshot.Style?.NodeTextColor ?? OfficeColor.White,
                OfficeTextAlignment.Center, verticalAlignment:
                OfficeTextVerticalAlignment.Center, wrapText: true,
                shrinkToFit: true);
        }
    }

    private readonly struct NodeBox {
        internal NodeBox(double x, double y, double width, double height) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double CenterX => X + Width / 2D;
        internal double CenterY => Y + Height / 2D;
    }
}
