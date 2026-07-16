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
    public static OfficeDrawing Render(OfficeDiagramSnapshot snapshot) {
        if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
        var drawing = new OfficeDrawing(snapshot.WidthPoints,
            snapshot.HeightPoints);
        drawing.AddShape(new OfficeShape {
            Kind = OfficeShapeKind.Rectangle,
            Width = snapshot.WidthPoints,
            Height = snapshot.HeightPoints,
            FillColor = OfficeColor.White
        }, 0D, 0D);

        List<NodeBox> nodes = snapshot.Kind switch {
            OfficeDiagramKind.Hierarchy => LayoutHierarchy(snapshot),
            OfficeDiagramKind.Cycle => LayoutCycle(snapshot),
            _ => LayoutProcess(snapshot)
        };
        AddConnectors(drawing, snapshot.Kind, nodes);
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
        int levels = checked((int)Math.Floor(Math.Log(count, 2D)) + 1);
        double levelHeight = snapshot.HeightPoints / levels;
        var result = new List<NodeBox>(count);
        for (int index = 0; index < count; index++) {
            int level = checked((int)Math.Floor(Math.Log(index + 1, 2D)));
            int firstIndex = (1 << level) - 1;
            int nodesOnLevel = Math.Min(1 << level, count - firstIndex);
            int position = index - firstIndex;
            double cellWidth = snapshot.WidthPoints / nodesOnLevel;
            double nodeWidth = Math.Min(Math.Max(12D, cellWidth * 0.62D),
                Math.Max(1D, cellWidth - 8D));
            double nodeHeight = Math.Min(Math.Max(10D, levelHeight * 0.48D),
                Math.Max(1D, levelHeight - 8D));
            result.Add(new NodeBox(
                position * cellWidth + (cellWidth - nodeWidth) / 2D,
                level * levelHeight + (levelHeight - nodeHeight) / 2D,
                nodeWidth, nodeHeight));
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

    private static void AddConnectors(OfficeDrawing drawing,
        OfficeDiagramKind kind, IReadOnlyList<NodeBox> nodes) {
        if (nodes.Count < 2) return;
        if (kind == OfficeDiagramKind.Hierarchy) {
            for (int index = 1; index < nodes.Count; index++) {
                AddConnector(drawing, nodes[(index - 1) / 2], nodes[index]);
            }
            return;
        }
        int connectorCount = kind == OfficeDiagramKind.Cycle
            ? nodes.Count
            : nodes.Count - 1;
        for (int index = 0; index < connectorCount; index++) {
            AddConnector(drawing, nodes[index],
                nodes[(index + 1) % nodes.Count]);
        }
    }

    private static void AddConnector(OfficeDrawing drawing, NodeBox source,
        NodeBox target) {
        double x1 = source.CenterX;
        double y1 = source.CenterY;
        double x2 = target.CenterX;
        double y2 = target.CenterY;
        if (Math.Abs(x1 - x2) < 0.000001D
            && Math.Abs(y1 - y2) < 0.000001D) return;
        OfficeShape line = OfficeShape.Line(x1, y1, x2, y2);
        line.StrokeColor = OfficeColor.FromRgb(100, 116, 139);
        line.StrokeWidth = 1.5D;
        line.StrokeEndMarker = new OfficeLineMarker(
            OfficeLineMarkerKind.Triangle, 5D, 5D);
        drawing.AddShape(line, Math.Min(x1, x2), Math.Min(y1, y2));
    }

    private static void AddNodes(OfficeDrawing drawing,
        OfficeDiagramSnapshot snapshot, IReadOnlyList<NodeBox> nodes) {
        for (int index = 0; index < nodes.Count; index++) {
            NodeBox node = nodes[index];
            OfficeShape shape = snapshot.Kind == OfficeDiagramKind.Cycle
                ? OfficeShape.Ellipse(node.Width, node.Height)
                : OfficeShape.RoundedRectangle(node.Width, node.Height,
                    Math.Min(8D, Math.Min(node.Width, node.Height) * 0.18D));
            shape.FillColor = NodeColors[index % NodeColors.Length];
            shape.StrokeColor = OfficeColor.White;
            shape.StrokeWidth = 1.25D;
            drawing.AddShape(shape, node.X, node.Y);
            double fontSize = Math.Max(6D, Math.Min(12D,
                Math.Min(node.Height * 0.28D, node.Width * 0.09D)));
            drawing.AddText(snapshot.Nodes[index], node.X + 3D,
                node.Y + 2D, Math.Max(1D, node.Width - 6D),
                Math.Max(1D, node.Height - 4D),
                new OfficeFontInfo("Calibri", fontSize,
                    OfficeFontStyle.Bold), OfficeColor.White,
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
