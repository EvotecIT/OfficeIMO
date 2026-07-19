using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private const int MaximumDrawingVisibilityDepth = 64;

    private static bool IsVisibleVisualPrimitive(
        PdfPageVisualPrimitive primitive,
        double pageWidth,
        double pageHeight,
        VisualGeometryBudget budget) {
        if (!HasFinitePrimitiveGeometry(primitive) ||
            !IsFinite(pageWidth) ||
            !IsFinite(pageHeight) ||
            pageWidth <= 0D ||
            pageHeight <= 0D) {
            return false;
        }

        bool hasVisibleFill = primitive.Kind != PdfPageVisualPrimitiveKind.Line &&
            ((HasOrdinaryFill(primitive) && HasVisibleOpacity(primitive.FillOpacity)) ||
             HasVisibleTilingPattern(primitive.FillTilingPattern, budget));
        bool hasVisibleStroke = primitive.StrokeWidth > 0D &&
            ((HasOrdinaryStroke(primitive) && HasVisibleOpacity(primitive.StrokeOpacity)) ||
             HasVisibleTilingPattern(primitive.StrokeTilingPattern, budget));
        if (!hasVisibleFill && !hasVisibleStroke) {
            return false;
        }

        PdfPageClipPath pageClip = PdfPageClipPath.Rectangle(0D, 0D, pageWidth, pageHeight);
        PdfPageClipPath visibleClip = pageClip;
        if (primitive.ClipPath.HasValue) {
            PdfPageClipPath authoredClip = primitive.ClipPath.Value;
            if (!HasFiniteClipGeometry(authoredClip)) {
                return false;
            }

            visibleClip = PdfPageClipPath.ResolveActiveClip(pageClip, authoredClip);
        }

        if (!HasFiniteClipGeometry(visibleClip) ||
            visibleClip.Width <= 0D ||
            visibleClip.Height <= 0D) {
            return false;
        }

        VisualPath? clipPath = VisualPath.FromClip(visibleClip, budget);
        if (clipPath == null) {
            return true;
        }

        if (hasVisibleFill) {
            VisualPath? fillPath = VisualPath.FromFill(primitive, budget);
            if (fillPath == null || fillPath.IntersectsFill(clipPath, budget)) {
                return true;
            }
        }

        if (hasVisibleStroke) {
            VisualPath? strokePath = VisualPath.FromStroke(primitive, budget);
            if (strokePath == null ||
                strokePath.StrokeIntersectsFill(clipPath, primitive.StrokeWidth / 2D, budget)) {
                return true;
            }
        }

        return budget.Exceeded;
    }

    private static bool HasFinitePrimitiveGeometry(PdfPageVisualPrimitive primitive) {
        if (!IsFinite(primitive.X) ||
            !IsFinite(primitive.Y) ||
            !IsFinite(primitive.Width) ||
            !IsFinite(primitive.Height) ||
            !IsFinite(primitive.X1) ||
            !IsFinite(primitive.Y1) ||
            !IsFinite(primitive.X2) ||
            !IsFinite(primitive.Y2) ||
            !IsFinite(primitive.StrokeWidth)) {
            return false;
        }

        for (int i = 0; i < primitive.PathCommands.Count; i++) {
            if (!HasFiniteCommand(primitive.PathCommands[i])) {
                return false;
            }
        }

        return true;
    }

    private static bool HasFiniteClipGeometry(PdfPageClipPath clip) {
        if (!IsFinite(clip.X) ||
            !IsFinite(clip.Y) ||
            !IsFinite(clip.Width) ||
            !IsFinite(clip.Height)) {
            return false;
        }

        for (int i = 0; i < clip.Commands.Count; i++) {
            if (!HasFiniteCommand(clip.Commands[i])) {
                return false;
            }
        }

        return true;
    }

    private static bool HasFiniteCommand(OfficePathCommand command) =>
        IsFinite(command.Point.X) &&
        IsFinite(command.Point.Y) &&
        IsFinite(command.ControlPoint1.X) &&
        IsFinite(command.ControlPoint1.Y) &&
        IsFinite(command.ControlPoint2.X) &&
        IsFinite(command.ControlPoint2.Y);

    private static bool HasOrdinaryFill(PdfPageVisualPrimitive primitive) =>
        HasVisibleColor(primitive.FillColor) ||
        HasVisibleGradient(primitive.FillGradient) ||
        HasVisibleGradient(primitive.FillRadialGradient);

    private static bool HasOrdinaryStroke(PdfPageVisualPrimitive primitive) =>
        HasVisibleColor(primitive.StrokeColor) ||
        HasVisibleGradient(primitive.StrokeGradient) ||
        HasVisibleGradient(primitive.StrokeRadialGradient);

    private static bool HasVisibleTilingPattern(
        PdfPageTilingPatternPaint? pattern,
        VisualGeometryBudget budget) =>
        pattern != null &&
        IsFinite(pattern.Opacity) &&
        pattern.Opacity > 0D &&
        (!pattern.Tint.HasValue || pattern.Tint.Value.A > 0) &&
        HasVisibleDrawingContent(pattern.Resource.Tile, budget);

    private static bool HasVisibleDrawingContent(
        OfficeDrawing drawing,
        VisualGeometryBudget budget) {
        VisualPath? canvas = VisualPath.Rectangle(
            0D,
            0D,
            drawing.Width,
            drawing.Height,
            OfficeTransform.Identity,
            budget);
        if (canvas == null) {
            return drawing.Elements.Count > 0;
        }

        return HasVisibleDrawingContent(
            drawing,
            OfficeTransform.Identity,
            new[] { canvas },
            budget,
            depth: 0);
    }

    private static bool HasVisibleDrawingContent(
        OfficeDrawing drawing,
        OfficeTransform transform,
        IReadOnlyList<VisualPath> clips,
        VisualGeometryBudget budget,
        int depth) {
        if (depth > MaximumDrawingVisibilityDepth) {
            budget.Exhaust();
            return drawing.Elements.Count > 0;
        }

        for (int i = 0; i < drawing.Elements.Count; i++) {
            OfficeDrawingElement element = drawing.Elements[i];
            switch (element) {
                case OfficeDrawingShape drawingShape:
                    if (HasVisibleDrawingShape(drawingShape, transform, clips, budget)) {
                        return true;
                    }
                    break;
                case OfficeDrawingImage image:
                    if (IsFinite(image.Opacity) &&
                        image.Opacity > 0D &&
                        IsVisibleRectangle(
                            0D,
                            0D,
                            1D,
                            1D,
                            image.Projection.CreateUnitSquareTransform().Then(transform),
                            clips,
                            budget)) {
                        return true;
                    }
                    break;
                case OfficeDrawingImagePattern imagePattern:
                    if (IsFinite(imagePattern.Opacity) &&
                        imagePattern.Opacity > 0D &&
                        IsVisiblePlacement(imagePattern.Layout.Area, transform, clips, budget)) {
                        return true;
                    }
                    break;
                case OfficeDrawingText text:
                    if (!string.IsNullOrEmpty(text.Text) &&
                        (!text.Color.HasValue || text.Color.Value.A > 0) &&
                        IsVisibleTextFrame(
                            text.X,
                            text.Y,
                            text.Width,
                            text.Height,
                            text.CreateFrameTransform(),
                            transform,
                            clips,
                            budget)) {
                        return true;
                    }
                    break;
                case OfficeDrawingRichText richText:
                    if (HasVisibleRichTextPaint(richText) &&
                        IsVisibleTextFrame(
                            richText.X,
                            richText.Y,
                            richText.Width,
                            richText.Height,
                            richText.CreateFrameTransform(),
                            transform,
                            clips,
                            budget)) {
                        return true;
                    }
                    break;
                case OfficeDrawingGroup group:
                    if (HasVisibleDrawingGroup(group, transform, clips, budget, depth)) {
                        return true;
                    }
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    if (IsFinite(effectGroup.Opacity) &&
                        effectGroup.Opacity > 0D &&
                        HasVisibleSoftMask(effectGroup.SoftMask, budget) &&
                        HasVisibleDrawingContent(
                            effectGroup.InnerDrawing,
                            effectGroup.Transform.Then(transform),
                            clips,
                            budget,
                            depth + 1)) {
                        return true;
                    }
                    break;
                case OfficeDrawingTilingPattern tilingPattern:
                    if (HasVisibleNestedTilingPattern(
                            tilingPattern,
                            transform,
                            clips,
                            budget,
                            depth)) {
                        return true;
                    }
                    break;
            }
        }

        return false;
    }

    private static bool HasVisibleDrawingShape(
        OfficeDrawingShape drawingShape,
        OfficeTransform parentTransform,
        IReadOnlyList<VisualPath> clips,
        VisualGeometryBudget budget) {
        OfficeShape shape = drawingShape.Shape;
        bool fill = shape.Kind != OfficeShapeKind.Line && HasVisibleShapeFill(shape);
        bool stroke = shape.StrokeWidth > 0D && HasVisibleShapeStroke(shape);
        bool shadow = shape.Shadow != null &&
            shape.Shadow.Color.A > 0 &&
            shape.Shadow.Opacity > 0D;
        bool glow = shape.Glow != null &&
            shape.Glow.Color.A > 0 &&
            shape.Glow.Opacity > 0D &&
            shape.Glow.Radius > 0D;
        if (!fill && !stroke && !shadow && !glow) {
            return false;
        }

        OfficeTransform shapeTransform = (shape.Transform ?? OfficeTransform.Identity)
            .Then(OfficeTransform.Translate(drawingShape.X, drawingShape.Y))
            .Then(parentTransform);
        IReadOnlyList<VisualPath> effectiveClips = clips;
        if (shape.ClipPath != null) {
            VisualPath? shapeClip = VisualPath.FromOfficeClip(shape.ClipPath, shapeTransform, budget);
            if (shapeClip == null) {
                return true;
            }

            effectiveClips = AppendClip(clips, shapeClip);
            if (!VisualPath.HasPositiveAreaIntersection(effectiveClips, budget)) {
                return budget.Exceeded;
            }
        }

        VisualPath? geometry = VisualPath.FromShape(shape, shapeTransform, budget);
        if (geometry == null) {
            return true;
        }

        if ((fill || shadow) && geometry.IntersectsFills(effectiveClips, budget)) {
            return true;
        }

        if (stroke || glow) {
            double strokeWidth = stroke ? shape.StrokeWidth : 0D;
            if (glow) {
                strokeWidth = Math.Max(strokeWidth, shape.StrokeWidth + (shape.Glow!.Radius * 2D));
            }

            double transformedHalfWidth = strokeWidth *
                VisualPath.GetMaximumScale(shapeTransform) /
                2D;
            if (geometry.StrokeIntersectsFills(effectiveClips, transformedHalfWidth, budget)) {
                return true;
            }
        }

        return budget.Exceeded;
    }

    private static bool HasVisibleDrawingGroup(
        OfficeDrawingGroup group,
        OfficeTransform parentTransform,
        IReadOnlyList<VisualPath> clips,
        VisualGeometryBudget budget,
        int depth) {
        OfficeTransform groupTransform = OfficeTransform.Translate(group.X, group.Y);
        if (group.FrameTransform.HasValue && group.FrameTransform.Value.HasTransform) {
            groupTransform = groupTransform.Then(group.FrameTransform.Value.CreateDestinationTransform());
        }
        groupTransform = groupTransform.Then(parentTransform);

        VisualPath? groupClip = VisualPath.FromOfficeClip(group.ClipPath, groupTransform, budget);
        if (groupClip == null) {
            return true;
        }

        IReadOnlyList<VisualPath> effectiveClips = AppendClip(clips, groupClip);
        if (!VisualPath.HasPositiveAreaIntersection(effectiveClips, budget)) {
            return budget.Exceeded;
        }

        OfficeTransform contentTransform = OfficeTransform.Translate(
                group.ContentOffsetX,
                group.ContentOffsetY)
            .Then(groupTransform);
        return HasVisibleDrawingContent(
            group.InnerDrawing,
            contentTransform,
            effectiveClips,
            budget,
            depth + 1);
    }

    private static bool HasVisibleNestedTilingPattern(
        OfficeDrawingTilingPattern pattern,
        OfficeTransform parentTransform,
        IReadOnlyList<VisualPath> clips,
        VisualGeometryBudget budget,
        int depth) {
        if (!IsFinite(pattern.Opacity) || pattern.Opacity <= 0D) {
            return false;
        }

        VisualPath? area = VisualPath.Rectangle(
            pattern.Area.X,
            pattern.Area.Y,
            pattern.Area.Width,
            pattern.Area.Height,
            parentTransform,
            budget);
        if (area == null) {
            return true;
        }

        IReadOnlyList<VisualPath> effectiveClips = AppendClip(clips, area);
        if (!VisualPath.HasPositiveAreaIntersection(effectiveClips, budget)) {
            return budget.Exceeded;
        }

        IReadOnlyList<OfficeTransform> tileTransforms = pattern.GetTileTransforms(pattern.MaximumTileCount);
        for (int i = 0; i < tileTransforms.Count; i++) {
            if (HasVisibleDrawingContent(
                    pattern.InnerTile,
                    tileTransforms[i].Then(parentTransform),
                    effectiveClips,
                    budget,
                    depth + 1)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasVisibleSoftMask(
        OfficeDrawingSoftMask? softMask,
        VisualGeometryBudget budget) {
        if (softMask == null) {
            return true;
        }

        bool visibleBackdrop = softMask.Mode == OfficeSoftMaskMode.Alpha
            ? softMask.BackdropColor.A > 0
            : softMask.BackdropColor.A > 0 &&
              (softMask.BackdropColor.R > 0 ||
               softMask.BackdropColor.G > 0 ||
               softMask.BackdropColor.B > 0);
        if (visibleBackdrop) {
            return true;
        }

        VisualPath? canvas = VisualPath.Rectangle(
            0D,
            0D,
            softMask.InnerDrawing.Width,
            softMask.InnerDrawing.Height,
            OfficeTransform.Identity,
            budget);
        if (canvas == null) {
            return softMask.InnerDrawing.Elements.Count > 0;
        }

        return HasVisibleDrawingContent(
            softMask.InnerDrawing,
            softMask.Transform,
            new[] { canvas },
            budget,
            depth: 0);
    }

    private static bool IsVisiblePlacement(
        OfficeImagePlacement placement,
        OfficeTransform transform,
        IReadOnlyList<VisualPath> clips,
        VisualGeometryBudget budget) =>
        IsVisibleRectangle(
            placement.X,
            placement.Y,
            placement.Width,
            placement.Height,
            transform,
            clips,
            budget);

    private static bool IsVisibleTextFrame(
        double x,
        double y,
        double width,
        double height,
        OfficeImageFrameTransform frameTransform,
        OfficeTransform parentTransform,
        IReadOnlyList<VisualPath> clips,
        VisualGeometryBudget budget) {
        OfficeTransform transform = frameTransform.HasTransform
            ? frameTransform.CreateDestinationTransform().Then(parentTransform)
            : parentTransform;
        return IsVisibleRectangle(x, y, width, height, transform, clips, budget);
    }

    private static bool IsVisibleRectangle(
        double x,
        double y,
        double width,
        double height,
        OfficeTransform transform,
        IReadOnlyList<VisualPath> clips,
        VisualGeometryBudget budget) {
        VisualPath? rectangle = VisualPath.Rectangle(x, y, width, height, transform, budget);
        return rectangle == null || rectangle.IntersectsFills(clips, budget);
    }

    private static List<VisualPath> AppendClip(
        IReadOnlyList<VisualPath> clips,
        VisualPath clip) {
        var result = new List<VisualPath>(clips.Count + 1);
        for (int i = 0; i < clips.Count; i++) {
            result.Add(clips[i]);
        }
        result.Add(clip);
        return result;
    }

    private static bool HasVisibleShapeFill(OfficeShape shape) =>
        (HasVisibleColor(shape.FillColor) ||
         HasVisibleGradient(shape.FillGradient) ||
         HasVisibleGradient(shape.FillRadialGradient)) &&
        HasVisibleOpacity(shape.FillOpacity);

    private static bool HasVisibleShapeStroke(OfficeShape shape) =>
        (HasVisibleColor(shape.StrokeColor) ||
         HasVisibleGradient(shape.StrokeGradient) ||
         HasVisibleGradient(shape.StrokeRadialGradient)) &&
        HasVisibleOpacity(shape.StrokeOpacity);

    private static bool HasVisibleRichTextPaint(OfficeDrawingRichText richText) {
        for (int i = 0; i < richText.Runs.Count; i++) {
            OfficeRichTextRun run = richText.Runs[i];
            if (!string.IsNullOrEmpty(run.Text) &&
                (run.Color.A > 0 ||
                 (run.BackgroundColor.HasValue && run.BackgroundColor.Value.A > 0))) {
                return true;
            }
        }

        return false;
    }

    private static bool HasVisibleColor(OfficeColor? color) =>
        color.HasValue && color.Value.A > 0;

    private static bool HasVisibleGradient(OfficeLinearGradient? gradient) {
        if (gradient == null) {
            return false;
        }

        for (int i = 0; i < gradient.Stops.Count; i++) {
            if (gradient.Stops[i].Color.A > 0) {
                return true;
            }
        }

        return false;
    }

    private static bool HasVisibleGradient(OfficeRadialGradient? gradient) {
        if (gradient == null) {
            return false;
        }

        for (int i = 0; i < gradient.Stops.Count; i++) {
            if (gradient.Stops[i].Color.A > 0) {
                return true;
            }
        }

        return false;
    }

    private static bool HasVisibleOpacity(double? opacity) =>
        !opacity.HasValue || (IsFinite(opacity.Value) && opacity.Value > 0D);
}
