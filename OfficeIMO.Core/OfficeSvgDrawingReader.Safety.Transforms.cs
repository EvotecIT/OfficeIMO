using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryResolveSupportedRasterTransform(
        XElement element,
        OfficeTransform inherited,
        double viewX,
        double viewY,
        out OfficeTransform transform) {
        string? value = ReadRasterProjectedAttribute(element, "transform");
        if (string.IsNullOrWhiteSpace(value)) {
            transform = inherited;
            return true;
        }
        if (!OfficeSvgTransformParser.TryParse(value, out OfficeTransform parsed)) {
            transform = inherited;
            return false;
        }

        OfficeTransform normalized = OfficeTransform.Translate(viewX, viewY)
            .Then(parsed)
            .Then(OfficeTransform.Translate(-viewX, -viewY));
        transform = normalized.Then(inherited);
        return IsSupportedSvgTransform(transform);
    }

    private static bool TryResolveRasterPixelScales(
        XElement root,
        double viewWidth,
        double viewHeight,
        double viewportWidth,
        double viewportHeight,
        out double pixelScaleX,
        out double pixelScaleY) {
        pixelScaleX = viewportWidth / viewWidth;
        pixelScaleY = viewportHeight / viewHeight;
        if (Math.Abs(pixelScaleX - 1D) < 0.000001D && Math.Abs(pixelScaleY - 1D) < 0.000001D) return true;
        if (!TryParsePreserveAspectRatio(
                root.Attribute("preserveAspectRatio")?.Value,
                out SvgAspectAlignment alignment,
                out bool slice)) return false;
        if (alignment == SvgAspectAlignment.None) return true;
        double uniformScale = slice
            ? Math.Max(pixelScaleX, pixelScaleY)
            : Math.Min(pixelScaleX, pixelScaleY);
        pixelScaleX = uniformScale;
        pixelScaleY = uniformScale;
        return true;
    }

    private static bool TryReadRasterUsePlacement(XElement use, out double x, out double y) {
        x = y = 0D;
        return TryReadRasterLength(use, "x", 0D, out x)
            && TryReadRasterLength(use, "y", 0D, out y);
    }

    private static bool TryResolveRenderedUseTargetTransform(
        XElement use,
        XElement target,
        OfficeTransform placedTransform,
        out OfficeTransform targetTransform) {
        targetTransform = placedTransform;
        if (!target.Name.LocalName.Equals("symbol", StringComparison.OrdinalIgnoreCase)) return true;
        if (!TryParseNumberList(ReadRasterProjectedAttribute(target, "viewBox"), out IReadOnlyList<double> viewBox)
            || viewBox.Count != 4
            || viewBox[2] <= 0D
            || viewBox[3] <= 0D
            || !TryReadRasterUseOrTargetLength(use, target, "width", viewBox[2], out double width)
            || !TryReadRasterUseOrTargetLength(use, target, "height", viewBox[3], out double height)
            || width <= 0D
            || height <= 0D
            || !TryParsePreserveAspectRatio(
                ReadRasterProjectedAttribute(use, "preserveAspectRatio")
                    ?? ReadRasterProjectedAttribute(target, "preserveAspectRatio"),
                out SvgAspectAlignment alignment,
                out bool slice)) return false;
        targetTransform = OfficeTransform.Translate(-viewBox[0], -viewBox[1])
            .Then(ResolveViewportTransform(viewBox[2], viewBox[3], width, height, alignment, slice))
            .Then(placedTransform);
        return IsSupportedSvgTransform(targetTransform);
    }

    private static bool TryResolveNestedRasterViewportTransform(
        XElement element,
        OfficeTransform inherited,
        out OfficeTransform transform) {
        transform = inherited;
        if (!TryReadRasterLength(element, "x", 0D, out double x)
            || !TryReadRasterLength(element, "y", 0D, out double y)) return false;

        string? viewBoxText = ReadRasterProjectedAttribute(element, "viewBox");
        if (string.IsNullOrWhiteSpace(viewBoxText)) {
            transform = OfficeTransform.Translate(x, y).Then(inherited);
            return IsSupportedSvgTransform(transform);
        }
        if (!TryParseNumberList(viewBoxText, out IReadOnlyList<double> viewBox)
            || viewBox.Count != 4
            || viewBox[2] <= 0D
            || viewBox[3] <= 0D
            || !TryReadRequiredRasterLength(element, "width", out double width)
            || !TryReadRequiredRasterLength(element, "height", out double height)
            || width <= 0D
            || height <= 0D
            || !TryParsePreserveAspectRatio(
                ReadRasterProjectedAttribute(element, "preserveAspectRatio"),
                out SvgAspectAlignment alignment,
                out bool slice)) return false;

        transform = OfficeTransform.Translate(-viewBox[0], -viewBox[1])
            .Then(ResolveViewportTransform(viewBox[2], viewBox[3], width, height, alignment, slice))
            .Then(OfficeTransform.Translate(x, y))
            .Then(inherited);
        return IsSupportedSvgTransform(transform);
    }

    private static bool TryReadRequiredRasterLength(XElement element, string name, out double value) {
        value = 0D;
        string? text = ReadRasterProjectedAttribute(element, name);
        return !string.IsNullOrWhiteSpace(text)
            && OfficeImageReader.TryParseSvgLength(text, out value)
            && !double.IsNaN(value)
            && !double.IsInfinity(value);
    }

    private static bool TryReadRasterUseOrTargetLength(
        XElement use,
        XElement target,
        string name,
        double fallback,
        out double value) {
        string? text = ReadRasterProjectedAttribute(use, name) ?? ReadRasterProjectedAttribute(target, name);
        if (string.IsNullOrWhiteSpace(text)) {
            value = fallback;
            return true;
        }
        return OfficeImageReader.TryParseSvgLength(text, out value)
            && !double.IsNaN(value)
            && !double.IsInfinity(value);
    }

    private static bool TryReadRasterLength(XElement element, string name, double fallback, out double value) {
        string? text = ReadRasterProjectedAttribute(element, name);
        if (string.IsNullOrWhiteSpace(text)) {
            value = fallback;
            return true;
        }
        return (TrySvgLength(text, out value) || OfficeImageReader.TryParseSvgLength(text, out value))
            && !double.IsNaN(value)
            && !double.IsInfinity(value);
    }
}
