using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    private static void AddGroupShape(OfficeDrawing drawing, PowerPointGroupShape groupShape,
        List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping,
        A.ColorScheme? colorScheme, PowerPointImageExportOptions options,
        PowerPointEmbeddedImageReadBudget imageBudget, int groupDepth, CancellationToken token = default) {
        if (options.MaxGroupShapeDepth >= 0 && groupDepth >= options.MaxGroupShapeDepth) {
            diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                OfficeDiagnosticCodes.GroupDepthLimit, "Skipped nested PowerPoint group content because MaxGroupShapeDepth was reached.",
                DescribeShape(groupShape), OfficeConversionLossKind.Omission));
            return;
        }
        if (groupShape.OwnerSlide == null) {
            AddUnsupportedShapeDiagnostic(diagnostics, groupShape, "Skipped a PowerPoint group shape because its owning slide context could not be resolved.");
            return;
        }

        OfficeImageFrameTransform frameTransform = CreateGroupFrameTransform(groupShape, 0D, 0D, 0D, 0D);
        if (frameTransform.HasTransform) {
            if (!TryGetBounds(groupShape, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return;
            }

            var groupDrawing = new OfficeDrawing(Math.Max(1D, drawing.Width - left), Math.Max(1D, drawing.Height - top));
            groupDrawing.Fonts.AddRange(drawing.Fonts);
            PowerPointShapeBoundsMapping localChildMapping = CreateGroupLocalChildMapping(groupShape, mapping);
            AddSlideContent(groupShape.OwnerSlide.GetGroupChildren(groupShape), groupDrawing, diagnostics,
                localChildMapping, colorScheme, options, imageBudget, groupDepth + 1, token);

            try {
                OfficeImageFrameTransform groupFrameTransform = CreateGroupFrameTransform(groupShape, left, top, width, height);
                if (RequiresGroupClip(groupDrawing, width, height)) {
                    drawing.AddClippedDrawing(groupDrawing, left, top, OfficeClipPath.Rectangle(width, height), groupFrameTransform);
                } else {
                    drawing.AddDrawing(groupDrawing, left, top, groupFrameTransform);
                }
            } catch (ArgumentOutOfRangeException) {
                AddUnsupportedShapeDiagnostic(diagnostics, groupShape, "Skipped a PowerPoint group shape because its transformed Drawing content would exceed the slide canvas.");
            }

            return;
        }

        if (TryGetBounds(groupShape, drawing, diagnostics, mapping, out double clipLeft, out double clipTop, out double clipWidth, out double clipHeight)) {
            var groupDrawing = new OfficeDrawing(Math.Max(1D, drawing.Width - clipLeft), Math.Max(1D, drawing.Height - clipTop));
            groupDrawing.Fonts.AddRange(drawing.Fonts);
            PowerPointShapeBoundsMapping localChildMapping = CreateGroupLocalChildMapping(groupShape, mapping);
            AddSlideContent(groupShape.OwnerSlide.GetGroupChildren(groupShape), groupDrawing, diagnostics,
                localChildMapping, colorScheme, options, imageBudget, groupDepth + 1, token);
            if (RequiresGroupClip(groupDrawing, clipWidth, clipHeight)) {
                drawing.AddClippedDrawing(groupDrawing, clipLeft, clipTop, OfficeClipPath.Rectangle(clipWidth, clipHeight));
            } else {
                drawing.AddDrawing(groupDrawing, clipLeft, clipTop);
            }
            return;
        }

        PowerPointShapeBoundsMapping childMapping = CreateGroupChildMapping(groupShape, mapping);
        AddSlideContent(groupShape.OwnerSlide.GetGroupChildren(groupShape), drawing, diagnostics,
            childMapping, colorScheme, options, imageBudget, groupDepth + 1, token);
    }

    private static bool RequiresGroupClip(OfficeDrawing groupDrawing, double width, double height) {
        for (int i = 0; i < groupDrawing.Elements.Count; i++) {
            OfficeDrawingElement element = groupDrawing.Elements[i];
            if (element is OfficeDrawingShape shape) {
                if (shape.X < 0D || shape.Y < 0D || shape.X + shape.Shape.Width > width || shape.Y + shape.Shape.Height > height) {
                    return true;
                }
            } else if (element is OfficeDrawingText text) {
                if (text.X < 0D || text.Y < 0D || text.X + text.Width > width || text.Y + text.Height > height) {
                    return true;
                }
            } else if (element is OfficeDrawingRichText richText) {
                if (richText.X < 0D || richText.Y < 0D || richText.X + richText.Width > width || richText.Y + richText.Height > height) {
                    return true;
                }
            } else if (element is OfficeDrawingImage image) {
                (double left, double top, double right, double bottom) = image.Projection.GetDestinationBounds();
                if (left < 0D || top < 0D || right > width || bottom > height) {
                    return true;
                }
            } else if (element is OfficeDrawingGroup group) {
                if (group.X < 0D || group.Y < 0D || group.X + group.ClipPath.Width > width || group.Y + group.ClipPath.Height > height) {
                    return true;
                }
            }
        }

        return false;
    }

    private static void ApplyShapeTransform(OfficeShape target, PowerPointShape source, double width, double height) {
        double rotation = source.Rotation ?? 0D;
        bool flipHorizontal = source.HorizontalFlip == true;
        bool flipVertical = source.VerticalFlip == true;
        var transform = new OfficeImageFrameTransform(
            rotation,
            width / 2D,
            height / 2D,
            flipHorizontal,
            flipVertical);
        if (transform.HasTransform) {
            target.Transform = transform.CreateDestinationTransform();
        }
    }

    private static bool TryGetBounds(PowerPointShape shape, OfficeDrawing drawing, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, out double left, out double top, out double width, out double height) {
        if (!shape.TryGetExportBoundsPoints(out left, out top, out width, out height)) {
            diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                OfficeDiagnosticCodes.InvalidShapeBounds,
                "Skipped a PowerPoint shape because its bounds are outside the slide drawing canvas.",
                DescribeShape(shape), OfficeConversionLossKind.Omission));
            return false;
        }

        left = mapping.MapX(left);
        top = mapping.MapY(top);
        width = mapping.MapWidth(width);
        height = mapping.MapHeight(height);

        if (width <= 0D ||
            height <= 0D ||
            left < 0D ||
            top < 0D ||
            left + width > drawing.Width ||
            top + height > drawing.Height) {
            diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                OfficeDiagnosticCodes.InvalidShapeBounds,
                "Skipped a PowerPoint shape because its bounds are outside the slide drawing canvas.",
                DescribeShape(shape), OfficeConversionLossKind.Omission));
            return false;
        }

        return true;
    }

    private static PowerPointShapeBoundsMapping CreateGroupChildMapping(PowerPointGroupShape groupShape, PowerPointShapeBoundsMapping parentMapping) {
        A.TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
        long? groupXEmu = transform?.Offset?.X?.Value;
        long? groupYEmu = transform?.Offset?.Y?.Value;
        long? groupWidthEmu = transform?.Extents?.Cx?.Value;
        long? groupHeightEmu = transform?.Extents?.Cy?.Value;
        long? childXEmu = transform?.ChildOffset?.X?.Value;
        long? childYEmu = transform?.ChildOffset?.Y?.Value;
        long? childWidthEmu = transform?.ChildExtents?.Cx?.Value;
        long? childHeightEmu = transform?.ChildExtents?.Cy?.Value;
        if (!groupXEmu.HasValue || !groupYEmu.HasValue || !groupWidthEmu.HasValue || !groupHeightEmu.HasValue ||
            !childXEmu.HasValue || !childYEmu.HasValue || !childWidthEmu.HasValue || !childHeightEmu.HasValue ||
            childWidthEmu.Value == 0L || childHeightEmu.Value == 0L) {
            return parentMapping;
        }

        double groupX = PowerPointUnits.ToPoints(groupXEmu.Value);
        double groupY = PowerPointUnits.ToPoints(groupYEmu.Value);
        double childX = PowerPointUnits.ToPoints(childXEmu.Value);
        double childY = PowerPointUnits.ToPoints(childYEmu.Value);
        double scaleX = groupWidthEmu.Value / (double)childWidthEmu.Value;
        double scaleY = groupHeightEmu.Value / (double)childHeightEmu.Value;
        PowerPointShapeBoundsMapping groupMapping = new PowerPointShapeBoundsMapping(
            groupX - (childX * scaleX),
            groupY - (childY * scaleY),
            scaleX,
            scaleY);
        return parentMapping.Compose(groupMapping);
    }

    private static PowerPointShapeBoundsMapping CreateGroupLocalChildMapping(PowerPointGroupShape groupShape, PowerPointShapeBoundsMapping parentMapping) {
        A.TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
        bool hasTransformedAncestor = parentMapping.HasTransformedAncestor
            || CreateGroupFrameTransform(groupShape, 0D, 0D, 0D, 0D).HasTransform;
        long? groupWidthEmu = transform?.Extents?.Cx?.Value;
        long? groupHeightEmu = transform?.Extents?.Cy?.Value;
        long? childXEmu = transform?.ChildOffset?.X?.Value;
        long? childYEmu = transform?.ChildOffset?.Y?.Value;
        long? childWidthEmu = transform?.ChildExtents?.Cx?.Value;
        long? childHeightEmu = transform?.ChildExtents?.Cy?.Value;
        if (!groupWidthEmu.HasValue || !groupHeightEmu.HasValue ||
            !childXEmu.HasValue || !childYEmu.HasValue || !childWidthEmu.HasValue || !childHeightEmu.HasValue ||
            childWidthEmu.Value == 0L || childHeightEmu.Value == 0L) {
            return parentMapping.WithTransformedAncestor(hasTransformedAncestor);
        }

        double childX = PowerPointUnits.ToPoints(childXEmu.Value);
        double childY = PowerPointUnits.ToPoints(childYEmu.Value);
        double scaleX = groupWidthEmu.Value / (double)childWidthEmu.Value;
        double scaleY = groupHeightEmu.Value / (double)childHeightEmu.Value;
        return new PowerPointShapeBoundsMapping(
            -parentMapping.MapWidth(childX * scaleX),
            -parentMapping.MapHeight(childY * scaleY),
            parentMapping.MapWidth(scaleX),
            parentMapping.MapHeight(scaleY),
            hasTransformedAncestor);
    }

    private static OfficeImageFrameTransform CreateGroupFrameTransform(PowerPointGroupShape groupShape, double left, double top, double width, double height) {
        A.TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
        double rotation = transform?.Rotation?.Value / 60000D ?? 0D;
        return new OfficeImageFrameTransform(
            rotation,
            left + (width / 2D),
            top + (height / 2D),
            transform?.HorizontalFlip?.Value == true,
            transform?.VerticalFlip?.Value == true);
    }

    private static bool HasUnsupportedTransform(PowerPointShape shape) {
        double rotation = shape.Rotation ?? 0D;
        return Math.Abs(rotation) > 0.000001D || shape.HorizontalFlip == true || shape.VerticalFlip == true;
    }
}
