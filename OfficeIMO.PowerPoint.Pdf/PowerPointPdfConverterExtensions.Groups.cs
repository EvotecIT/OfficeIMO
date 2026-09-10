using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderGroupShape(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointGroupShape groupShape,
        double groupX, double groupY, double groupWidth, double groupHeight,
        int slideNumber, double pageWidth, double pageHeight, PowerPointToPdfOptions options, bool warnInvalidBounds, int groupDepth) {
        if (options.MaxGroupShapeDepth >= 0 && groupDepth >= options.MaxGroupShapeDepth) {
            AddWarning(options, slideNumber, "group-depth-limit", "Skipped nested PowerPoint group shape content because MaxGroupShapeDepth was reached.");
            return;
        }

        TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
        var frame = new OfficeImageFrameTransform(transform?.Rotation?.Value / 60000D ?? 0D,
            groupX + groupWidth / 2D, groupY + groupHeight / 2D,
            transform?.HorizontalFlip?.Value == true, transform?.VerticalFlip?.Value == true);
        if (frame.HasTransform) {
            canvas.Effect(frame.CreateDestinationTransform(), 1D, target =>
                RenderGroupChildren(target, groupShape, groupX, groupY, groupWidth, groupHeight,
                    slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth));
        } else {
            RenderGroupChildren(canvas, groupShape, groupX, groupY, groupWidth, groupHeight,
                slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth);
        }
    }

    private static void RenderGroupChildren(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointGroupShape groupShape,
        double groupX, double groupY, double groupWidth, double groupHeight,
        int slideNumber, double pageWidth, double pageHeight, PowerPointToPdfOptions options, bool warnInvalidBounds, int groupDepth) {
        IReadOnlyList<PptCore.PowerPointShape> children = groupShape.OwnerSlide!.GetGroupChildren(groupShape);
        foreach (PptCore.PowerPointShape child in children) {
            if (child.Hidden) {
                continue;
            }

            child.TryGetExportBoundsPoints(out double x, out double y, out double width, out double height);
            bool renderable = IsLineShape(child) ? width >= 0 && height >= 0 && (width > 0 || height > 0)
                : width > 0 && height > 0;
            if (!renderable) {
                if (warnInvalidBounds) AddWarning(options, slideNumber, "invalid-shape-bounds", "Skipped a PowerPoint group child with non-positive PDF bounds.");
                continue;
            }

            MapGroupChildBox(groupShape, groupX, groupY, groupWidth, groupHeight, ref x, ref y, ref width, ref height);
            // The root group owns the page-space clip. A local clip here would rotate with
            // the ancestors and discard children that transform back onto the slide.
            RenderShapeContent(canvas, child, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth + 1);
        }
    }

    private static void MapGroupChildBox(PptCore.PowerPointGroupShape groupShape,
        double groupX, double groupY, double groupWidth, double groupHeight,
        ref double x, ref double y, ref double width, ref double height) {
        TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
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
            return;
        }

        double childX = PptCore.PowerPointUnits.ToPoints(childXEmu.Value);
        double childY = PptCore.PowerPointUnits.ToPoints(childYEmu.Value);
        double scaleX = groupWidth / PptCore.PowerPointUnits.ToPoints(childWidthEmu.Value);
        double scaleY = groupHeight / PptCore.PowerPointUnits.ToPoints(childHeightEmu.Value);
        x = groupX + (x - childX) * scaleX;
        y = groupY + (y - childY) * scaleY;
        width *= scaleX;
        height *= scaleY;
    }

}
