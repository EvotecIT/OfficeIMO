using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private const int MaximumPatternVisibilityTiles = 16384;

    private static bool HasVisibleTilingPattern(
        PdfPageTilingPatternPaint pattern,
        VisualPath paintedGeometry,
        double strokeHalfWidth,
        IReadOnlyList<VisualPath> clips,
        Dictionary<PdfPageTilingPatternResource, bool> patternPaintCache,
        VisualGeometryBudget budget) {
        if (!patternPaintCache.TryGetValue(pattern.Resource, out bool tileMayPaint)) {
            bool completed = TryDrawingMayPaint(
                pattern.Resource.Tile,
                budget,
                depth: 0,
                out tileMayPaint);
            if (completed) {
                patternPaintCache[pattern.Resource] = tileMayPaint;
            }
        }

        if (!tileMayPaint) {
            return false;
        }
        if (budget.Exceeded) {
            return true;
        }

        VisualPath? paintedArea = paintedGeometry;
        if (strokeHalfWidth > 0D) {
            VisualBounds bounds = paintedGeometry.GetBounds(strokeHalfWidth);
            paintedArea = VisualPath.Rectangle(
                bounds.Left,
                bounds.Top,
                bounds.Width,
                bounds.Height,
                OfficeTransform.Identity,
                budget);
        }
        if (paintedArea == null) {
            return true;
        }

        IReadOnlyList<VisualPath> effectiveClips = AppendClip(clips, paintedArea);
        if (!VisualPath.HasPositiveAreaIntersection(effectiveClips, budget)) {
            return budget.Exceeded;
        }

        return HasVisibleRepeatedDrawing(
            pattern.Resource.Tile,
            pattern.Resource.HorizontalStep,
            pattern.Resource.VerticalStep,
            pattern.Transform,
            originX: 0D,
            originY: 0D,
            effectiveClips,
            MaximumPatternVisibilityTiles,
            budget,
            depth: 0);
    }

    private static bool HasVisibleRepeatedDrawing(
        OfficeDrawing tile,
        double horizontalStep,
        double verticalStep,
        OfficeTransform patternTransform,
        double originX,
        double originY,
        IReadOnlyList<VisualPath> clips,
        int maximumTileCount,
        VisualGeometryBudget budget,
        int depth) {
        if (!IsFinite(horizontalStep) ||
            !IsFinite(verticalStep) ||
            horizontalStep <= 0D ||
            verticalStep <= 0D ||
            !IsFinite(tile.Width) ||
            !IsFinite(tile.Height) ||
            tile.Width <= 0D ||
            tile.Height <= 0D ||
            maximumTileCount <= 0 ||
            !VisualPath.TryGetCommonBounds(clips, out VisualBounds targetBounds)) {
            return false;
        }
        if (!patternTransform.TryInvert(out OfficeTransform inverse)) {
            budget.Exhaust();
            return true;
        }

        OfficePoint topLeft = inverse.TransformPoint(new OfficePoint(targetBounds.Left, targetBounds.Top));
        OfficePoint topRight = inverse.TransformPoint(new OfficePoint(targetBounds.Right, targetBounds.Top));
        OfficePoint bottomRight = inverse.TransformPoint(new OfficePoint(targetBounds.Right, targetBounds.Bottom));
        OfficePoint bottomLeft = inverse.TransformPoint(new OfficePoint(targetBounds.Left, targetBounds.Bottom));
        double minX = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomRight.X, bottomLeft.X));
        double maxX = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomRight.X, bottomLeft.X));
        double minY = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomRight.Y, bottomLeft.Y));
        double maxY = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomRight.Y, bottomLeft.Y));
        if (!TryGetTileRange(
                minX,
                maxX,
                originX,
                tile.Width,
                horizontalStep,
                out long firstColumn,
                out long lastColumn) ||
            !TryGetTileRange(
                minY,
                maxY,
                originY,
                tile.Height,
                verticalStep,
                out long firstRow,
                out long lastRow)) {
            budget.Exhaust();
            return true;
        }

        double columnCount = (double)lastColumn - firstColumn + 1D;
        double rowCount = (double)lastRow - firstRow + 1D;
        if (columnCount <= 0D || rowCount <= 0D) {
            return false;
        }
        if (!IsFinite(columnCount) ||
            !IsFinite(rowCount) ||
            columnCount > maximumTileCount ||
            rowCount > maximumTileCount ||
            columnCount * rowCount > maximumTileCount) {
            budget.Exhaust();
            return true;
        }
        int columns = (int)columnCount;
        int rows = (int)rowCount;

        for (int rowOffset = 0; rowOffset < rows; rowOffset++) {
            long row = firstRow + rowOffset;
            for (int columnOffset = 0; columnOffset < columns; columnOffset++) {
                long column = firstColumn + columnOffset;
                if (!budget.TryUseOperation()) {
                    return true;
                }

                double tileX = originX + (column * horizontalStep);
                double tileY = originY + (row * verticalStep);
                if (!IsFinite(tileX) || !IsFinite(tileY)) {
                    budget.Exhaust();
                    return true;
                }
                OfficeTransform tileTransform = OfficeTransform.Translate(
                        tileX,
                        tileY)
                    .Then(patternTransform);
                if (HasVisibleDrawingContent(
                        tile,
                        tileTransform,
                        clips,
                        budget,
                        depth + 1)) {
                    return true;
                }
            }
        }

        return budget.Exceeded;
    }

    private static bool TryGetTileRange(
        double minimum,
        double maximum,
        double origin,
        double tileSize,
        double step,
        out long first,
        out long last) {
        first = 0L;
        last = -1L;
        double firstValue = Math.Floor((minimum - origin - tileSize) / step) + 1D;
        double lastValue = Math.Ceiling((maximum - origin) / step) - 1D;
        const double MinimumLong = -9223372036854775808D;
        const double MaximumLong = 9223372036854774784D;
        if (!IsFinite(firstValue) ||
            !IsFinite(lastValue) ||
            firstValue < MinimumLong ||
            firstValue > MaximumLong ||
            lastValue < MinimumLong ||
            lastValue > MaximumLong) {
            return false;
        }

        first = (long)firstValue;
        last = (long)lastValue;
        return true;
    }

    private static bool TryDrawingMayPaint(
        OfficeDrawing drawing,
        VisualGeometryBudget budget,
        int depth,
        out bool mayPaint) {
        mayPaint = false;
        if (depth > MaximumDrawingVisibilityDepth) {
            budget.Exhaust();
            mayPaint = drawing.Elements.Count > 0;
            return false;
        }

        for (int i = 0; i < drawing.Elements.Count; i++) {
            if (!budget.TryUseOperation()) {
                mayPaint = true;
                return false;
            }

            switch (drawing.Elements[i]) {
                case OfficeDrawingShape drawingShape:
                    OfficeShape shape = drawingShape.Shape;
                    if ((shape.Kind != OfficeShapeKind.Line && HasVisibleShapeFill(shape)) ||
                        (shape.StrokeWidth > 0D && HasVisibleShapeStroke(shape)) ||
                        (shape.Shadow != null &&
                         shape.Shadow.Color.A > 0 &&
                         shape.Shadow.Opacity > 0D) ||
                        (shape.Glow != null &&
                         shape.Glow.Color.A > 0 &&
                         shape.Glow.Opacity > 0D &&
                         shape.Glow.Radius > 0D)) {
                        mayPaint = true;
                        return true;
                    }
                    break;
                case OfficeDrawingImage image:
                    if (IsFinite(image.Opacity) && image.Opacity > 0D) {
                        mayPaint = true;
                        return true;
                    }
                    break;
                case OfficeDrawingImagePattern imagePattern:
                    if (IsFinite(imagePattern.Opacity) && imagePattern.Opacity > 0D) {
                        mayPaint = true;
                        return true;
                    }
                    break;
                case OfficeDrawingText text:
                    if (!string.IsNullOrEmpty(text.Text) &&
                        (!text.Color.HasValue || text.Color.Value.A > 0)) {
                        mayPaint = true;
                        return true;
                    }
                    break;
                case OfficeDrawingRichText richText:
                    if (HasVisibleRichTextPaint(richText)) {
                        mayPaint = true;
                        return true;
                    }
                    break;
                case OfficeDrawingGroup group:
                    if (!TryDrawingMayPaint(
                            group.InnerDrawing,
                            budget,
                            depth + 1,
                            out bool groupMayPaint)) {
                        mayPaint = groupMayPaint;
                        return false;
                    }
                    if (groupMayPaint) {
                        mayPaint = true;
                        return true;
                    }
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    if (IsFinite(effectGroup.Opacity) &&
                        effectGroup.Opacity > 0D) {
                        if (!TryDrawingMayPaint(
                                effectGroup.InnerDrawing,
                                budget,
                                depth + 1,
                                out bool effectMayPaint)) {
                            mayPaint = effectMayPaint;
                            return false;
                        }
                        if (effectMayPaint) {
                            mayPaint = true;
                            return true;
                        }
                    }
                    break;
                case OfficeDrawingTilingPattern pattern:
                    if (IsFinite(pattern.Opacity) &&
                        pattern.Opacity > 0D) {
                        if (!TryDrawingMayPaint(
                                pattern.InnerTile,
                                budget,
                                depth + 1,
                                out bool patternMayPaint)) {
                            mayPaint = patternMayPaint;
                            return false;
                        }
                        if (patternMayPaint) {
                            mayPaint = true;
                            return true;
                        }
                    }
                    break;
            }
        }

        return true;
    }
}
