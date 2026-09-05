using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static void TryAddSvgLink(XElement element, OfficeDrawing linkedContent, OfficeDrawing target, ref int unsupported) {
        XAttribute[] hrefAttributes = element.Attributes()
            .Where(attribute => attribute.Name.LocalName.Equals("href", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (hrefAttributes.Length != 1 || string.IsNullOrWhiteSpace(hrefAttributes[0].Value)) {
            unsupported++;
            return;
        }
        if (!TryGetSvgDrawingBounds(linkedContent.Elements, out SvgInteractiveBounds bounds)) return;
        double left = Math.Max(0D, bounds.Left);
        double top = Math.Max(0D, bounds.Top);
        double right = Math.Min(target.Width, bounds.Right);
        double bottom = Math.Min(target.Height, bounds.Bottom);
        if (right <= left || bottom <= top) return;
        string? alternativeText = element.Attributes()
            .FirstOrDefault(attribute => attribute.Name.LocalName.Equals("aria-label", StringComparison.OrdinalIgnoreCase))?.Value;
        if (string.IsNullOrWhiteSpace(alternativeText)) {
            alternativeText = element.Elements()
                .FirstOrDefault(child => child.Name.LocalName.Equals("title", StringComparison.OrdinalIgnoreCase))?.Value;
        }
        target.AddLink(hrefAttributes[0].Value, left, top, right - left, bottom - top, alternativeText);
    }

    private static bool TryGetSvgDrawingBounds(IEnumerable<OfficeDrawingElement> elements, out SvgInteractiveBounds bounds) {
        bounds = default;
        bool hasBounds = false;
        foreach (OfficeDrawingElement element in elements) {
            if (!TryGetSvgElementBounds(element, out SvgInteractiveBounds current)) continue;
            bounds = hasBounds ? bounds.Union(current) : current;
            hasBounds = true;
        }
        return hasBounds;
    }

    private static bool TryGetSvgElementBounds(OfficeDrawingElement element, out SvgInteractiveBounds bounds) {
        if (element is OfficeDrawingShape drawingShape) {
            bounds = new SvgInteractiveBounds(drawingShape.X, drawingShape.Y,
                drawingShape.X + drawingShape.Shape.Width, drawingShape.Y + drawingShape.Shape.Height);
            if (drawingShape.Shape.Transform.HasValue) bounds = bounds.Transform(drawingShape.Shape.Transform.Value);
            return true;
        }
        if (element is OfficeDrawingText text) {
            (double left, double top, double right, double bottom) = OfficeGeometry.GetRotatedRectangleBounds(
                text.X, text.Y, text.Width, text.Height, text.RotationDegrees, text.RotationCenterX, text.RotationCenterY);
            bounds = new SvgInteractiveBounds(left, top, right, bottom);
            return true;
        }
        if (element is OfficeDrawingRichText richText) {
            (double left, double top, double right, double bottom) = OfficeGeometry.GetRotatedRectangleBounds(
                richText.X, richText.Y, richText.Width, richText.Height, richText.RotationDegrees,
                richText.RotationCenterX, richText.RotationCenterY);
            bounds = new SvgInteractiveBounds(left, top, right, bottom);
            return true;
        }
        if (element is OfficeDrawingImage image) {
            (double left, double top, double right, double bottom) = image.Projection.GetDestinationBounds();
            bounds = new SvgInteractiveBounds(left, top, right, bottom);
            return true;
        }
        if (element is OfficeDrawingImagePattern imagePattern) {
            OfficeImagePlacement area = imagePattern.Layout.Area;
            bounds = new SvgInteractiveBounds(area.X, area.Y, area.X + area.Width, area.Y + area.Height);
            return true;
        }
        if (element is OfficeDrawingTilingPattern pattern) {
            OfficeImagePlacement area = pattern.Area;
            bounds = new SvgInteractiveBounds(area.X, area.Y, area.X + area.Width, area.Y + area.Height);
            return true;
        }
        if (element is OfficeDrawingGroup group) {
            bounds = new SvgInteractiveBounds(group.X, group.Y,
                group.X + group.ClipPath.Width, group.Y + group.ClipPath.Height);
            if (group.FrameTransform.HasValue && group.FrameTransform.Value.HasTransform) {
                bounds = bounds.Transform(group.FrameTransform.Value.CreateDestinationTransform());
            }
            return true;
        }
        if (element is OfficeDrawingEffectGroup effectGroup
            && TryGetSvgDrawingBounds(effectGroup.InnerDrawing.Elements, out SvgInteractiveBounds effectBounds)) {
            bounds = effectBounds.Transform(effectGroup.Transform);
            return true;
        }
        bounds = default;
        return false;
    }

    private readonly struct SvgInteractiveBounds {
        internal SvgInteractiveBounds(double left, double top, double right, double bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }
        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal SvgInteractiveBounds Union(SvgInteractiveBounds other) => new SvgInteractiveBounds(
            Math.Min(Left, other.Left), Math.Min(Top, other.Top), Math.Max(Right, other.Right), Math.Max(Bottom, other.Bottom));
        internal SvgInteractiveBounds Transform(OfficeTransform transform) {
            (double left, double top, double right, double bottom) = transform.TransformRectangleBounds(
                Left, Top, Right - Left, Bottom - Top);
            return new SvgInteractiveBounds(left, top, right, bottom);
        }
    }
}
