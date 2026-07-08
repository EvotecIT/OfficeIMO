using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double SmartArtFallbackNodeHeight = 28D;
        private const double SmartArtFallbackNodeGap = 8D;
        private const double SmartArtFallbackPadding = 8D;

        private static bool AddSmartArt(WordSmartArt smartArt, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            int nodeCount = smartArt.NodeCount;
            if (nodeCount == 0) {
                if (context.IsTargetPage) {
                    AddDiagnostic(diagnostics, "unsupported-word-smartart", "Skipped a Word SmartArt diagram because no editable node text could be read.", "Word SmartArt");
                }

                return false;
            }

            if (!TryGetSmartArtSize(smartArt, out double authoredWidth, out double authoredHeight)) {
                authoredWidth = context.ContentWidth;
                authoredHeight = SmartArtFallbackPadding * 2D + (nodeCount * SmartArtFallbackNodeHeight) + ((nodeCount - 1) * SmartArtFallbackNodeGap);
            }

            bool horizontalProcess = IsHorizontalProcessSmartArt(smartArt);
            bool cycle = IsCycleSmartArt(smartArt);
            bool hierarchy = IsHierarchySmartArt(smartArt);
            double width = Math.Min(Math.Max(cycle || hierarchy ? 160D : 120D, authoredWidth), context.ContentWidth);
            double nodeHeight = SmartArtFallbackNodeHeight;
            double neededHeight = horizontalProcess
                ? SmartArtFallbackPadding * 2D + nodeHeight
                : cycle
                    ? Math.Max(144D, SmartArtFallbackPadding * 2D + nodeHeight)
                    : hierarchy
                        ? Math.Max(112D, SmartArtFallbackPadding * 2D + (nodeHeight * 2D) + SmartArtFallbackNodeGap)
                : SmartArtFallbackPadding * 2D + (nodeCount * nodeHeight) + ((nodeCount - 1) * SmartArtFallbackNodeGap);
            double height = Math.Max(Math.Min(Math.Max(48D, authoredHeight), context.ContentHeight), neededHeight);
            if (!EnsureVerticalSpace(context, height, diagnostics)) {
                return false;
            }

            if (context.IsTargetPage) {
                if (TryAddSmartArtPersistedLayoutDrawing(smartArt, context, width, height)) {
                    // Authored persisted diagram layout was projected through shared Drawing primitives.
                } else if (horizontalProcess && TryAddSmartArtHorizontalProcessDrawing(smartArt, nodeCount, context, width, height, nodeHeight)) {
                    AddDiagnostic(
                        diagnostics,
                        "limited-word-smartart",
                        "Rendered Word SmartArt Basic Process as a dependency-free process fallback because exact SmartArt layout is not implemented yet.",
                        "Word SmartArt");
                } else if (cycle && TryAddSmartArtCycleDrawing(smartArt, nodeCount, context, width, height, nodeHeight)) {
                    AddDiagnostic(
                        diagnostics,
                        "limited-word-smartart",
                        "Rendered Word SmartArt Cycle as a dependency-free circular fallback because exact SmartArt layout is not implemented yet.",
                        "Word SmartArt");
                } else if (hierarchy && TryAddSmartArtHierarchyDrawing(smartArt, nodeCount, context, width, height, nodeHeight)) {
                    AddDiagnostic(
                        diagnostics,
                        "limited-word-smartart",
                        "Rendered Word SmartArt hierarchy as a dependency-free organization fallback because exact SmartArt layout is not implemented yet.",
                        "Word SmartArt");
                } else {
                    AddSmartArtFallbackDrawing(smartArt, nodeCount, context, width, height, nodeHeight);
                    AddDiagnostic(
                        diagnostics,
                        "limited-word-smartart",
                        "Rendered Word SmartArt as a dependency-free text fallback because exact SmartArt layout is not implemented yet.",
                        "Word SmartArt");
                }
            }

            context.Y += height + ParagraphGapPoints;
            return true;
        }

        private static bool IsHorizontalProcessSmartArt(WordSmartArt smartArt) =>
            smartArt.LayoutType == SmartArtType.BasicProcess || smartArt.LayoutType == SmartArtType.ContinuousBlockProcess;

        private static bool IsCycleSmartArt(WordSmartArt smartArt) =>
            smartArt.LayoutType == SmartArtType.Cycle;

        private static bool IsHierarchySmartArt(WordSmartArt smartArt) =>
            smartArt.LayoutType == SmartArtType.Hierarchy || smartArt.LayoutType == SmartArtType.PictureOrgChart;

        private static bool TryAddSmartArtPersistedLayoutDrawing(WordSmartArt smartArt, WordImageFlowContext context, double width, double height) {
            IReadOnlyList<WordSmartArtPersistedShape> persistedShapes = smartArt.GetPersistedLayoutShapes();
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
                double shapeWidth = Math.Max(1D, persistedShape.Width * scaleX);
                double shapeHeight = Math.Max(1D, persistedShape.Height * scaleY);
                if (!OfficeShapePresets.TryCreate(persistedShape.PresetName, shapeWidth, shapeHeight, out OfficeShape? drawingShape) || drawingShape == null) {
                    return false;
                }
            }

            for (int i = 0; i < persistedShapes.Count; i++) {
                WordSmartArtPersistedShape persistedShape = persistedShapes[i];
                double shapeWidth = Math.Max(1D, persistedShape.Width * scaleX);
                double shapeHeight = Math.Max(1D, persistedShape.Height * scaleY);
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

                double left = context.Left + ((persistedShape.X - minX) * scaleX);
                double top = context.Y + ((persistedShape.Y - minY) * scaleY);
                context.Drawing.AddShape(drawingShape, left, top);
                if (!string.IsNullOrWhiteSpace(persistedShape.Text)) {
                    context.Drawing.AddText(
                        persistedShape.Text,
                        left + 4D,
                        top + 4D,
                        Math.Max(1D, shapeWidth - 8D),
                        Math.Max(1D, shapeHeight - 8D),
                        new OfficeFontInfo("Calibri", Math.Max(7D, Math.Min(10D, shapeHeight * 0.28D)), OfficeFontStyle.Bold),
                        OfficeColor.White,
                        OfficeTextAlignment.Center,
                        Math.Max(9D, Math.Min(13D, shapeHeight * 0.34D)),
                        wrapText: true);
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

        private static bool TryAddSmartArtHorizontalProcessDrawing(WordSmartArt smartArt, int nodeCount, WordImageFlowContext context, double width, double height, double nodeHeight) {
            double availableWidth = Math.Max(1D, width - (SmartArtFallbackPadding * 2D));
            double connectorGap = nodeCount <= 1 ? 0D : Math.Min(28D, Math.Max(10D, availableWidth * 0.06D));
            double totalConnectorWidth = connectorGap * Math.Max(0, nodeCount - 1);
            double nodeWidth = (availableWidth - totalConnectorWidth) / nodeCount;
            if (nodeWidth < 36D) {
                return false;
            }

            AddSmartArtFrame(context, width, height);

            double resolvedNodeHeight = Math.Max(nodeHeight, Math.Min(48D, height - (SmartArtFallbackPadding * 2D)));
            double nodeTop = context.Y + Math.Max(SmartArtFallbackPadding, (height - resolvedNodeHeight) / 2D);
            double nodeLeft = context.Left + SmartArtFallbackPadding;
            double connectorY = nodeTop + resolvedNodeHeight / 2D;
            for (int i = 0; i < nodeCount; i++) {
                string text = smartArt.GetNodeText(i);
                AddSmartArtFallbackNode(context, nodeLeft, nodeTop, nodeWidth, resolvedNodeHeight, text);
                if (i < nodeCount - 1 && connectorGap > 0D) {
                    AddSmartArtProcessConnector(context, nodeLeft + nodeWidth + 2D, connectorY, Math.Max(1D, connectorGap - 4D));
                }

                nodeLeft += nodeWidth + connectorGap;
            }

            return true;
        }

        private static bool TryAddSmartArtHierarchyDrawing(WordSmartArt smartArt, int nodeCount, WordImageFlowContext context, double width, double height, double nodeHeight) {
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
            double rootLeft = context.Left + SmartArtFallbackPadding + (availableWidth - rootWidth) / 2D;
            double rootTop = context.Y + SmartArtFallbackPadding;
            double childTop = childCount == 0
                ? rootTop
                : Math.Min(context.Y + height - SmartArtFallbackPadding - resolvedNodeHeight, rootTop + resolvedNodeHeight + verticalGap);

            AddSmartArtFrame(context, width, height);
            var root = new SmartArtNodeLayout(rootLeft, rootTop, rootWidth, resolvedNodeHeight, smartArt.GetNodeText(0));
            var children = new List<SmartArtNodeLayout>(childCount);
            double childLeft = context.Left + SmartArtFallbackPadding;
            for (int i = 0; i < childCount; i++) {
                children.Add(new SmartArtNodeLayout(childLeft, childTop, childWidth, resolvedNodeHeight, smartArt.GetNodeText(i + 1)));
                childLeft += childWidth + childGap;
            }

            for (int i = 0; i < children.Count; i++) {
                SmartArtNodeLayout child = children[i];
                AddSmartArtPlainConnector(context, root.CenterX, root.Y + root.Height, child.CenterX, child.Y);
            }

            AddSmartArtFallbackNode(context, root.X, root.Y, root.Width, root.Height, root.Text);
            for (int i = 0; i < children.Count; i++) {
                SmartArtNodeLayout child = children[i];
                AddSmartArtFallbackNode(context, child.X, child.Y, child.Width, child.Height, child.Text);
            }

            return true;
        }

        private static bool TryAddSmartArtCycleDrawing(WordSmartArt smartArt, int nodeCount, WordImageFlowContext context, double width, double height, double nodeHeight) {
            double availableWidth = Math.Max(1D, width - (SmartArtFallbackPadding * 2D));
            double availableHeight = Math.Max(1D, height - (SmartArtFallbackPadding * 2D));
            double nodeWidth = Math.Min(96D, Math.Max(48D, availableWidth * 0.34D));
            double resolvedNodeHeight = Math.Max(nodeHeight, Math.Min(40D, availableHeight * 0.24D));
            if (availableWidth < nodeWidth || availableHeight < resolvedNodeHeight) {
                return false;
            }

            AddSmartArtFrame(context, width, height);

            double centerX = context.Left + width / 2D;
            double centerY = context.Y + height / 2D;
            double radiusX = Math.Max(0D, (availableWidth - nodeWidth) / 2D);
            double radiusY = Math.Max(0D, (availableHeight - resolvedNodeHeight) / 2D);
            var nodes = new List<SmartArtNodeLayout>(nodeCount);
            for (int i = 0; i < nodeCount; i++) {
                double angle = (-Math.PI / 2D) + (2D * Math.PI * i / nodeCount);
                double left = Clamp(
                    centerX + Math.Cos(angle) * radiusX - nodeWidth / 2D,
                    context.Left + SmartArtFallbackPadding,
                    context.Left + width - SmartArtFallbackPadding - nodeWidth);
                double top = Clamp(
                    centerY + Math.Sin(angle) * radiusY - resolvedNodeHeight / 2D,
                    context.Y + SmartArtFallbackPadding,
                    context.Y + height - SmartArtFallbackPadding - resolvedNodeHeight);
                nodes.Add(new SmartArtNodeLayout(left, top, nodeWidth, resolvedNodeHeight, smartArt.GetNodeText(i)));
            }

            if (nodeCount > 1) {
                for (int i = 0; i < nodes.Count; i++) {
                    SmartArtNodeLayout current = nodes[i];
                    SmartArtNodeLayout next = nodes[(i + 1) % nodes.Count];
                    AddSmartArtConnector(context, current.CenterX, current.CenterY, next.CenterX, next.CenterY);
                }
            }

            for (int i = 0; i < nodes.Count; i++) {
                SmartArtNodeLayout node = nodes[i];
                AddSmartArtFallbackNode(context, node.X, node.Y, node.Width, node.Height, node.Text);
            }

            return true;
        }

        private static void AddSmartArtFrame(WordImageFlowContext context, double width, double height) {
            OfficeShape frame = OfficeShape.RoundedRectangle(width, height, 8D);
            frame.FillColor = OfficeColor.FromRgb(241, 245, 249);
            frame.StrokeColor = OfficeColor.FromRgb(148, 163, 184);
            frame.StrokeWidth = 1D;
            context.Drawing.AddShape(frame, context.Left, context.Y);
        }

        private static void AddSmartArtProcessConnector(WordImageFlowContext context, double x, double y, double length) {
            OfficeShape connector = OfficeShape.Line(0D, 0D, length, 0D);
            connector.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
            connector.StrokeWidth = 1.4D;
            connector.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 6D, 7D);
            context.Drawing.AddShape(connector, x, y);
        }

        private static void AddSmartArtPlainConnector(WordImageFlowContext context, double x1, double y1, double x2, double y2) {
            if (Math.Abs(x1 - x2) < 0.01D && Math.Abs(y1 - y2) < 0.01D) {
                return;
            }

            double left = Math.Min(x1, x2);
            double top = Math.Min(y1, y2);
            OfficeShape connector = OfficeShape.Line(new OfficePoint(x1, y1), new OfficePoint(x2, y2));
            connector.StrokeColor = OfficeColor.FromRgb(71, 85, 105);
            connector.StrokeWidth = 1.2D;
            context.Drawing.AddShape(connector, left, top);
        }

        private static void AddSmartArtConnector(WordImageFlowContext context, double x1, double y1, double x2, double y2) {
            if (Math.Abs(x1 - x2) < 0.01D && Math.Abs(y1 - y2) < 0.01D) {
                return;
            }

            double left = Math.Min(x1, x2);
            double top = Math.Min(y1, y2);
            OfficeShape connector = OfficeShape.Line(new OfficePoint(x1, y1), new OfficePoint(x2, y2));
            connector.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
            connector.StrokeWidth = 1.2D;
            connector.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 5D, 6D);
            context.Drawing.AddShape(connector, left, top);
        }

        private static void AddSmartArtFallbackDrawing(WordSmartArt smartArt, int nodeCount, WordImageFlowContext context, double width, double height, double nodeHeight) {
            AddSmartArtFrame(context, width, height);
            double availableHeight = Math.Max(nodeHeight, height - (SmartArtFallbackPadding * 2D));
            double gap = nodeCount <= 1
                ? 0D
                : Math.Min(SmartArtFallbackNodeGap, Math.Max(2D, (availableHeight - (nodeCount * nodeHeight)) / (nodeCount - 1)));
            double currentY = context.Y + SmartArtFallbackPadding;
            double nodeWidth = Math.Max(24D, width - (SmartArtFallbackPadding * 2D));
            for (int i = 0; i < nodeCount; i++) {
                string text = smartArt.GetNodeText(i);
                AddSmartArtFallbackNode(context, context.Left + SmartArtFallbackPadding, currentY, nodeWidth, nodeHeight, text);
                currentY += nodeHeight + gap;
            }
        }

        private static void AddSmartArtFallbackNode(WordImageFlowContext context, double x, double y, double width, double height, string text) {
            OfficeShape shape = OfficeShape.RoundedRectangle(width, height, 6D);
            shape.FillColor = OfficeColor.FromRgb(219, 234, 254);
            shape.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
            shape.StrokeWidth = 1D;
            context.Drawing.AddShape(shape, x, y);
            context.Drawing.AddText(
                text,
                x + 6D,
                y + 5D,
                Math.Max(1D, width - 12D),
                Math.Max(1D, height - 10D),
                new OfficeFontInfo("Calibri", 10D, OfficeFontStyle.Bold),
                OfficeColor.FromRgb(15, 23, 42),
                OfficeTextAlignment.Center,
                12D,
                wrapText: true);
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
}
