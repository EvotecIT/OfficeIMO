using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Projects a cached shape text frame and its margin inset into page coordinates without editing the model.</summary>
internal readonly struct VisioTextFramePlacement {
    private VisioTextFramePlacement(double x, double y, double width, double height, double angle) {
        PageX = x; PageY = y; ContentWidth = width; ContentHeight = height; Angle = angle;
    }

    internal double PageX { get; }
    internal double PageY { get; }
    internal double ContentWidth { get; }
    internal double ContentHeight { get; }
    internal double Angle { get; }

    /// <summary>
    /// Rotates the native local-pin offset and asymmetric margins in text-frame axes,
    /// then applies each containing shape's existing transform. Renderer fitting minimums
    /// are deliberately separate from the cached frame's position.
    /// </summary>
    internal static VisioTextFramePlacement Resolve(VisioShape shape, double drawingToPhysical, VisioNativeShapeTransform? nativeTransform = null) {
        VisioNativeShapeTransform transform = nativeTransform ?? VisioNativeShapeTransform.Create(shape);
        VisioTextStyle? style = shape.TextStyle;
        double width = style?.TextWidth ?? shape.Width;
        double height = style?.TextHeight ?? shape.Height;
        // Text-block margins are physical inches even on a scaled drawing. Convert them
        // before composing the inset with the cached drawing-unit frame and local pins.
        double left = (style?.LeftMargin ?? 0.05D) / drawingToPhysical, right = (style?.RightMargin ?? 0.05D) / drawingToPhysical;
        double top = (style?.TopMargin ?? 0.03D) / drawingToPhysical, bottom = (style?.BottomMargin ?? 0.03D) / drawingToPhysical;
        double angle = style?.TextAngle ?? 0D;
        // Visio's local Y axis points up: a larger top margin moves the content down.
        var offset = new OfficePoint(width / 2D - (style?.TextLocPinX ?? width / 2D) + (left - right) / 2D,
            height / 2D - (style?.TextLocPinY ?? height / 2D) + (bottom - top) / 2D);
        // MS-VSDX 2.2.8.5.1: subtract the text local pin, reflect in the shape's
        // local axes, rotate by TxtAngle, and finally add TxtPin before going to page.
        OfficePoint rotated = OfficeTransform.Scale(transform.FlipX ? -1 : 1, transform.FlipY ? -1 : 1)
            .Then(OfficeTransform.RotateDegrees(OfficeGeometry.RadiansToDegrees(angle))).TransformPoint(offset);
        double x = (style?.TextPinX ?? shape.Width / 2D) + rotated.X;
        double y = (style?.TextPinY ?? shape.Height / 2D) + rotated.Y;
        OfficePoint center = transform.PagePoint(x, y);
        return new VisioTextFramePlacement(center.X, center.Y, width - left - right, height - top - bottom, transform.TextAngle(angle));
    }
}
