using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Paints supported SmartArt semantics into a local drawing; placement and pagination belong to the consumer.</summary>
internal static class WordSmartArtDrawingRenderer {
    private const double SmartArtFallbackNodeHeight = 28D;
    private const double SmartArtFallbackNodeGap = 8D;
    private const double SmartArtFallbackPadding = 8D;

    /// <summary>
    /// Projects persisted preset geometry when it represents all editable node text, otherwise paints a named approximate layout.
    /// Persisted painting retains positions and presets, with simplified colors and typography.
    /// </summary>
    internal static OfficeDrawing? CreateDrawing(WordSmartArt smartArt, out string projection, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        int nodeCount = smartArt.NodeCount;
        var texts = new List<string>(nodeCount);
        for (int index = 0; index < nodeCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            texts.Add(smartArt.GetNodeText(index));
        }
        IReadOnlyList<WordSmartArtPersistedShape> persisted = smartArt.GetPersistedLayoutShapes();
        projection = "unavailable";
        if (nodeCount == 0 && persisted.Count == 0) return null;

        bool process = IsHorizontalProcessSmartArt(smartArt);
        bool cycle = IsCycleSmartArt(smartArt);
        bool hierarchy = IsHierarchySmartArt(smartArt);
        bool authoredSize = TryGetSmartArtSize(smartArt, out double width, out double height);
        if (!authoredSize) {
            width = 432D;
            height = SmartArtFallbackPadding * 2D + nodeCount * SmartArtFallbackNodeHeight +
                Math.Max(0, nodeCount - 1) * SmartArtFallbackNodeGap;
        }
        width = Math.Max(cycle || hierarchy ? 160D : 120D, width);
        height = Math.Max(48D, height);
        var drawing = new OfficeDrawing(width, height);
        if (HasAllNodeText(persisted, texts) && TryAddSmartArtPersistedLayoutDrawing(persisted, drawing, width, height)) {
            projection = "persisted";
            return drawing;
        }
        if (nodeCount == 0) return null;

        double neededHeight = process ? SmartArtFallbackPadding * 2D + SmartArtFallbackNodeHeight
            : cycle ? 144D : hierarchy ? 112D
            : SmartArtFallbackPadding * 2D + nodeCount * SmartArtFallbackNodeHeight +
                Math.Max(0, nodeCount - 1) * SmartArtFallbackNodeGap;
        height = Math.Max(height, neededHeight);
        drawing = new OfficeDrawing(width, height);
        if (process && TryAddSmartArtHorizontalProcessDrawing(texts, nodeCount, drawing, width, height, SmartArtFallbackNodeHeight))
            projection = "process";
        else if (cycle && TryAddSmartArtCycleDrawing(texts, nodeCount, drawing, width, height, SmartArtFallbackNodeHeight))
            projection = "cycle";
        else if (hierarchy && TryAddSmartArtHierarchyDrawing(texts, nodeCount, drawing, width, height, SmartArtFallbackNodeHeight))
            projection = "hierarchy";
        else {
            height = Math.Max(height, SmartArtFallbackPadding * 2D + nodeCount * SmartArtFallbackNodeHeight +
                Math.Max(0, nodeCount - 1) * SmartArtFallbackNodeGap);
            drawing = new OfficeDrawing(width, height);
            AddSmartArtFallbackDrawing(texts, nodeCount, drawing, width, height, SmartArtFallbackNodeHeight);
            projection = "text";
        }
        cancellationToken.ThrowIfCancellationRequested();
        return drawing;
    }

    private static bool HasAllNodeText(IReadOnlyList<WordSmartArtPersistedShape> shapes, IReadOnlyList<string> texts) {
        var available = shapes.Where(shape => !string.IsNullOrWhiteSpace(shape.Text))
            .GroupBy(shape => shape.Text, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        foreach (string text in texts) {
            if (string.IsNullOrWhiteSpace(text)) continue;
            if (!available.TryGetValue(text, out int count) || count == 0) return false;
            available[text] = count - 1;
        }
        return true;
    }

    private static bool IsHorizontalProcessSmartArt(WordSmartArt smartArt) =>
        smartArt.LayoutType == WordSmartArtType.BasicProcess || smartArt.LayoutType == WordSmartArtType.ContinuousBlockProcess;

    private static bool IsCycleSmartArt(WordSmartArt smartArt) =>
        smartArt.LayoutType == WordSmartArtType.Cycle;

    private static bool IsHierarchySmartArt(WordSmartArt smartArt) =>
        smartArt.LayoutType == WordSmartArtType.Hierarchy || smartArt.LayoutType == WordSmartArtType.PictureOrgChart;

    private static bool TryAddSmartArtPersistedLayoutDrawing(IReadOnlyList<WordSmartArtPersistedShape> persistedShapes, OfficeDrawing drawing, double width, double height) {
        if (persistedShapes.Count == 0) {
            return false;
        }

        double minX = persistedShapes.Min(shape => shape.X);
        double minY = persistedShapes.Min(shape => shape.Y);
        double maxX = persistedShapes.Max(shape => shape.X + shape.Width);
        double maxY = persistedShapes.Max(shape => shape.Y + shape.Height);
        double sourceWidth = maxX - minX;
        double sourceHeight = maxY - minY;
        if (sourceWidth <= 0D || sourceHeight <= 0D) {
            return false;
        }

        double scaleX = width / sourceWidth;
        double scaleY = height / sourceHeight;
        for (int i = 0; i < persistedShapes.Count; i++) {
            WordSmartArtPersistedShape persistedShape = persistedShapes[i];
            double left = Math.Max(0D, (persistedShape.X - minX) * scaleX);
            double top = Math.Max(0D, (persistedShape.Y - minY) * scaleY);
            // Normalize rounding at the far edge before the strict drawing bounds check.
            double shapeWidth = Math.Min(Math.Max(1D, persistedShape.Width * scaleX), width - left);
            double shapeHeight = Math.Min(Math.Max(1D, persistedShape.Height * scaleY), height - top);
            if (!OfficeShapePresets.TryCreate(persistedShape.PresetName, shapeWidth, shapeHeight, out OfficeShape? drawingShape) || drawingShape == null) {
                return false;
            }
        }

        for (int i = 0; i < persistedShapes.Count; i++) {
            WordSmartArtPersistedShape persistedShape = persistedShapes[i];
            double left = Math.Max(0D, (persistedShape.X - minX) * scaleX);
            double top = Math.Max(0D, (persistedShape.Y - minY) * scaleY);
            // Normalize rounding at the far edge before the strict drawing bounds check.
            double shapeWidth = Math.Min(Math.Max(1D, persistedShape.Width * scaleX), width - left);
            double shapeHeight = Math.Min(Math.Max(1D, persistedShape.Height * scaleY), height - top);
            OfficeShapePresets.TryCreate(persistedShape.PresetName, shapeWidth, shapeHeight, out OfficeShape? drawingShape);
            if (drawingShape == null) {
                continue;
            }

            ApplyPersistedSmartArtShapeStyle(drawingShape, persistedShape);
            if (Math.Abs(persistedShape.RotationDegrees) > 0.000001D) {
                drawingShape.Transform = OfficeTransform.Translate(-shapeWidth / 2D, -shapeHeight / 2D)
                    .Then(OfficeTransform.RotateDegrees(persistedShape.RotationDegrees))
                    .Then(OfficeTransform.Translate(shapeWidth / 2D, shapeHeight / 2D));
            }

            drawing.AddShape(drawingShape, left, top);
            if (!string.IsNullOrWhiteSpace(persistedShape.Text)) {
                drawing.AddText(
                    persistedShape.Text,
                    left + 4D,
                    top + 4D,
                    Math.Max(1D, shapeWidth - 8D),
                    Math.Max(1D, shapeHeight - 8D),
                    new OfficeFontInfo("Calibri", Math.Max(7D, Math.Min(10D, shapeHeight * 0.28D)), OfficeFontStyle.Bold),
                    OfficeColor.White,
                    OfficeTextAlignment.Center,
                    Math.Max(9D, Math.Min(13D, shapeHeight * 0.34D)),
                    wrapText: true, shrinkToFit: true);
            }
        }

        return true;
    }

    private static void ApplyPersistedSmartArtShapeStyle(OfficeShape shape, WordSmartArtPersistedShape persistedShape) {
        if (string.Equals(persistedShape.PresetName, "rightArrow", StringComparison.OrdinalIgnoreCase)) {
            shape.FillColor = OfficeColor.FromRgb(191, 219, 254);
            shape.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
            shape.StrokeWidth = 1D;
            return;
        }

        shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
        shape.StrokeColor = OfficeColor.White;
        shape.StrokeWidth = 1D;
    }

    private static bool TryAddSmartArtHorizontalProcessDrawing(IReadOnlyList<string> texts, int nodeCount, OfficeDrawing drawing, double width, double height, double nodeHeight) {
        double availableWidth = Math.Max(1D, width - (SmartArtFallbackPadding * 2D));
        double connectorGap = nodeCount <= 1 ? 0D : Math.Min(28D, Math.Max(10D, availableWidth * 0.06D));
        double totalConnectorWidth = connectorGap * Math.Max(0, nodeCount - 1);
        double nodeWidth = (availableWidth - totalConnectorWidth) / nodeCount;
        if (nodeWidth < 36D) {
            return false;
        }

        AddSmartArtFrame(drawing, width, height);

        double resolvedNodeHeight = Math.Max(nodeHeight, Math.Min(48D, height - (SmartArtFallbackPadding * 2D)));
        double nodeTop = Math.Max(SmartArtFallbackPadding, (height - resolvedNodeHeight) / 2D);
        double nodeLeft = SmartArtFallbackPadding;
        double connectorY = nodeTop + resolvedNodeHeight / 2D;
        for (int i = 0; i < nodeCount; i++) {
            string text = texts[i];
            AddSmartArtFallbackNode(drawing, nodeLeft, nodeTop, nodeWidth, resolvedNodeHeight, text);
            if (i < nodeCount - 1 && connectorGap > 0D) {
                AddSmartArtProcessConnector(drawing, nodeLeft + nodeWidth + 2D, connectorY, Math.Max(1D, connectorGap - 4D));
            }

            nodeLeft += nodeWidth + connectorGap;
        }

        return true;
    }

    private static bool TryAddSmartArtHierarchyDrawing(IReadOnlyList<string> texts, int nodeCount, OfficeDrawing drawing, double width, double height, double nodeHeight) {
        double availableWidth = Math.Max(1D, width - (SmartArtFallbackPadding * 2D));
        double availableHeight = Math.Max(1D, height - (SmartArtFallbackPadding * 2D));
        double resolvedNodeHeight = Math.Max(nodeHeight, Math.Min(42D, availableHeight * 0.32D));
        double verticalGap = Math.Min(36D, Math.Max(SmartArtFallbackNodeGap, availableHeight - (resolvedNodeHeight * 2D)));
        int childCount = Math.Max(0, nodeCount - 1);
        double childGap = childCount <= 1 ? 0D : Math.Min(18D, Math.Max(8D, availableWidth * 0.05D));
        double childWidth = childCount == 0
            ? Math.Min(120D, Math.Max(54D, availableWidth * 0.45D))
            : (availableWidth - (childGap * (childCount - 1))) / childCount;
        if (childWidth < 36D || availableHeight < resolvedNodeHeight) {
            return false;
        }

        double rootWidth = Math.Min(Math.Max(childWidth, 72D), Math.Min(140D, availableWidth));
        double rootLeft = SmartArtFallbackPadding + (availableWidth - rootWidth) / 2D;
        double rootTop = SmartArtFallbackPadding;
        double childTop = childCount == 0
            ? rootTop
            : Math.Min(height - SmartArtFallbackPadding - resolvedNodeHeight, rootTop + resolvedNodeHeight + verticalGap);

        AddSmartArtFrame(drawing, width, height);
        var root = new SmartArtNodeLayout(rootLeft, rootTop, rootWidth, resolvedNodeHeight, texts[0]);
        var children = new List<SmartArtNodeLayout>(childCount);
        double childLeft = SmartArtFallbackPadding;
        for (int i = 0; i < childCount; i++) {
            children.Add(new SmartArtNodeLayout(childLeft, childTop, childWidth, resolvedNodeHeight, texts[i + 1]));
            childLeft += childWidth + childGap;
        }

        for (int i = 0; i < children.Count; i++) {
            SmartArtNodeLayout child = children[i];
            AddSmartArtPlainConnector(drawing, root.CenterX, root.Y + root.Height, child.CenterX, child.Y);
        }

        AddSmartArtFallbackNode(drawing, root.X, root.Y, root.Width, root.Height, root.Text);
        for (int i = 0; i < children.Count; i++) {
            SmartArtNodeLayout child = children[i];
            AddSmartArtFallbackNode(drawing, child.X, child.Y, child.Width, child.Height, child.Text);
        }

        return true;
    }

    private static bool TryAddSmartArtCycleDrawing(IReadOnlyList<string> texts, int nodeCount, OfficeDrawing drawing, double width, double height, double nodeHeight) {
        double availableWidth = Math.Max(1D, width - (SmartArtFallbackPadding * 2D));
        double availableHeight = Math.Max(1D, height - (SmartArtFallbackPadding * 2D));
        double nodeWidth = Math.Min(96D, Math.Max(48D, availableWidth * 0.34D));
        double resolvedNodeHeight = Math.Max(nodeHeight, Math.Min(40D, availableHeight * 0.24D));
        if (availableWidth < nodeWidth || availableHeight < resolvedNodeHeight) {
            return false;
        }

        AddSmartArtFrame(drawing, width, height);

        double centerX = width / 2D;
        double centerY = height / 2D;
        double radiusX = Math.Max(0D, (availableWidth - nodeWidth) / 2D);
        double radiusY = Math.Max(0D, (availableHeight - resolvedNodeHeight) / 2D);
        var nodes = new List<SmartArtNodeLayout>(nodeCount);
        for (int i = 0; i < nodeCount; i++) {
            double angle = (-Math.PI / 2D) + (2D * Math.PI * i / nodeCount);
            double left = Clamp(
                centerX + Math.Cos(angle) * radiusX - nodeWidth / 2D,
                SmartArtFallbackPadding,
                width - SmartArtFallbackPadding - nodeWidth);
            double top = Clamp(
                centerY + Math.Sin(angle) * radiusY - resolvedNodeHeight / 2D,
                SmartArtFallbackPadding,
                height - SmartArtFallbackPadding - resolvedNodeHeight);
            nodes.Add(new SmartArtNodeLayout(left, top, nodeWidth, resolvedNodeHeight, texts[i]));
        }

        if (nodeCount > 1) {
            for (int i = 0; i < nodes.Count; i++) {
                SmartArtNodeLayout current = nodes[i];
                SmartArtNodeLayout next = nodes[(i + 1) % nodes.Count];
                AddSmartArtConnector(drawing, current.CenterX, current.CenterY, next.CenterX, next.CenterY);
            }
        }

        for (int i = 0; i < nodes.Count; i++) {
            SmartArtNodeLayout node = nodes[i];
            AddSmartArtFallbackNode(drawing, node.X, node.Y, node.Width, node.Height, node.Text);
        }

        return true;
    }

    private static void AddSmartArtFrame(OfficeDrawing drawing, double width, double height) {
        OfficeShape frame = OfficeShape.RoundedRectangle(width, height, 8D);
        frame.FillColor = OfficeColor.FromRgb(241, 245, 249);
        frame.StrokeColor = OfficeColor.FromRgb(148, 163, 184);
        frame.StrokeWidth = 1D;
        drawing.AddShape(frame, 0D, 0D);
    }

    private static void AddSmartArtProcessConnector(OfficeDrawing drawing, double x, double y, double length) {
        OfficeShape connector = OfficeShape.Line(0D, 0D, length, 0D);
        connector.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
        connector.StrokeWidth = 1.4D;
        connector.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 6D, 7D);
        drawing.AddShape(connector, x, y);
    }

    private static void AddSmartArtPlainConnector(OfficeDrawing drawing, double x1, double y1, double x2, double y2) {
        if (Math.Abs(x1 - x2) < 0.01D && Math.Abs(y1 - y2) < 0.01D) {
            return;
        }

        double left = Math.Min(x1, x2);
        double top = Math.Min(y1, y2);
        OfficeShape connector = OfficeShape.Line(new OfficePoint(x1, y1), new OfficePoint(x2, y2));
        connector.StrokeColor = OfficeColor.FromRgb(71, 85, 105);
        connector.StrokeWidth = 1.2D;
        drawing.AddShape(connector, left, top);
    }

    private static void AddSmartArtConnector(OfficeDrawing drawing, double x1, double y1, double x2, double y2) {
        if (Math.Abs(x1 - x2) < 0.01D && Math.Abs(y1 - y2) < 0.01D) {
            return;
        }

        double left = Math.Min(x1, x2);
        double top = Math.Min(y1, y2);
        OfficeShape connector = OfficeShape.Line(new OfficePoint(x1, y1), new OfficePoint(x2, y2));
        connector.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
        connector.StrokeWidth = 1.2D;
        connector.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 5D, 6D);
        drawing.AddShape(connector, left, top);
    }

    private static void AddSmartArtFallbackDrawing(IReadOnlyList<string> texts, int nodeCount, OfficeDrawing drawing, double width, double height, double nodeHeight) {
        AddSmartArtFrame(drawing, width, height);
        double availableHeight = Math.Max(nodeHeight, height - (SmartArtFallbackPadding * 2D));
        double gap = nodeCount <= 1
            ? 0D
            : Math.Min(SmartArtFallbackNodeGap, Math.Max(2D, (availableHeight - (nodeCount * nodeHeight)) / (nodeCount - 1)));
        double currentY = SmartArtFallbackPadding;
        double nodeWidth = Math.Max(24D, width - (SmartArtFallbackPadding * 2D));
        for (int i = 0; i < nodeCount; i++) {
            string text = texts[i];
            AddSmartArtFallbackNode(drawing, SmartArtFallbackPadding, currentY, nodeWidth, nodeHeight, text);
            currentY += nodeHeight + gap;
        }
    }

    private static void AddSmartArtFallbackNode(OfficeDrawing drawing, double x, double y, double width, double height, string text) {
        OfficeShape shape = OfficeShape.RoundedRectangle(width, height, 6D);
        shape.FillColor = OfficeColor.FromRgb(219, 234, 254);
        shape.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
        shape.StrokeWidth = 1D;
        drawing.AddShape(shape, x, y);
        drawing.AddText(
            text,
            x + 6D,
            y + 5D,
            Math.Max(1D, width - 12D),
            Math.Max(1D, height - 10D),
            new OfficeFontInfo("Calibri", 10D, OfficeFontStyle.Bold),
            OfficeColor.FromRgb(15, 23, 42),
            OfficeTextAlignment.Center,
            12D,
            wrapText: true, shrinkToFit: true);
    }

    private static bool TryGetSmartArtSize(WordSmartArt smartArt, out double width, out double height) {
        width = 0D;
        height = 0D;
        DW.Extent? extent = smartArt._drawing.GetFirstChild<DW.Inline>()?.Extent
            ?? smartArt._drawing.GetFirstChild<DW.Anchor>()?.Extent;
        if (extent?.Cx?.Value is long cx && extent.Cy?.Value is long cy && cx > 0L && cy > 0L) {
            width = Helpers.ConvertEmusToPoints(cx);
            height = Helpers.ConvertEmusToPoints(cy);
            return width > 0D && height > 0D;
        }

        return false;
    }

    private static double Clamp(double value, double minimum, double maximum) =>
        Math.Max(minimum, Math.Min(maximum, value));

    private readonly struct SmartArtNodeLayout {
        internal SmartArtNodeLayout(double x, double y, double width, double height, string text) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Text = text;
        }

        internal double X { get; }

        internal double Y { get; }

        internal double Width { get; }

        internal double Height { get; }

        internal string Text { get; }

        internal double CenterX => X + Width / 2D;

        internal double CenterY => Y + Height / 2D;
    }
}
