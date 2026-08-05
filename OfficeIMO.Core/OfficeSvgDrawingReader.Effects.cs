using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryResolveSvgEffects(
        XElement element,
        double width,
        double height,
        SvgPaintContext inheritedStyle,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform transform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref int unsupported,
        out OfficeBlendMode blendMode,
        out OfficeDrawingSoftMask? softMask) {
        blendMode = OfficeBlendMode.Normal;
        softMask = null;
        bool hasEffects = false;

        string? blendValue = ReadPresentationProperty(element, "mix-blend-mode");
        if (!string.IsNullOrWhiteSpace(blendValue)) {
            if (!TryParseBlendMode(blendValue!, out blendMode)) {
                hasEffects = true;
                unsupported++;
            } else if (blendMode != OfficeBlendMode.Normal) {
                hasEffects = true;
            }
        }

        string? maskValue = ReadPresentationProperty(element, "mask");
        if (string.IsNullOrWhiteSpace(maskValue) || maskValue!.Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) {
            return hasEffects;
        }

        hasEffects = true;
        if (!references.TryEnterLocal(maskValue, out string maskId, out XElement? maskElement)) {
            unsupported++;
            return hasEffects;
        }
        try {
            if (maskElement == null || !maskElement.Name.LocalName.Equals("mask", StringComparison.OrdinalIgnoreCase)) {
                unsupported++;
                return hasEffects;
            }
            string? maskUnits = maskElement.Attribute("maskUnits")?.Value;
            if (string.IsNullOrWhiteSpace(maskUnits)
                || !maskUnits!.Trim().Equals("userSpaceOnUse", StringComparison.OrdinalIgnoreCase)) {
                unsupported++;
                return hasEffects;
            }
            string? contentUnits = maskElement.Attribute("maskContentUnits")?.Value;
            if (!string.IsNullOrWhiteSpace(contentUnits)
                && !contentUnits!.Trim().Equals("userSpaceOnUse", StringComparison.OrdinalIgnoreCase)) {
                unsupported++;
                return hasEffects;
            }

            if (!TryResolveUserSpaceMaskRegion(maskElement, width, height, viewX, viewY,
                    out double regionX, out double regionY, out double regionWidth, out double regionHeight)) {
                unsupported++;
                return hasEffects;
            }

            var maskContent = new OfficeDrawing(width, height);
            SvgPaintContext maskStyle = ResolveDefinitionPaintContext(maskElement, paintServers, ref unsupported);
            OfficeTransform maskTransform = ResolveTransform(maskElement, transform, viewX, viewY, ref unsupported);
            AddChildren(maskElement, maskContent, maskStyle, paintServers, references, maskTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                ref visited, ref pathCommands, ref unsupported);
            OfficeDrawing maskDrawing = ClipMaskToRegion(maskContent, regionX, regionY, regionWidth, regionHeight, transform);
            OfficeSoftMaskMode mode = ReadPresentationProperty(maskElement, "mask-type")?.Trim().Equals("alpha", StringComparison.OrdinalIgnoreCase) == true
                ? OfficeSoftMaskMode.Alpha
                : OfficeSoftMaskMode.Luminosity;
            softMask = new OfficeDrawingSoftMask(maskDrawing, mode);
            return true;
        } finally {
            references.Exit(maskId);
        }
    }

    private static SvgPaintContext ResolveDefinitionPaintContext(
        XElement definition,
        SvgPaintServerRegistry paintServers,
        ref int unsupported) {
        SvgPaintContext context = SvgPaintContext.Default;
        foreach (XElement ancestor in definition.Ancestors().Reverse()) {
            context = ResolvePaintContext(ancestor, context, paintServers, ref unsupported);
        }
        return ResolvePaintContext(definition, context, paintServers, ref unsupported);
    }

    private static bool TryResolveUserSpaceMaskRegion(
        XElement maskElement,
        double width,
        double height,
        double viewX,
        double viewY,
        out double x,
        out double y,
        out double regionWidth,
        out double regionHeight) {
        x = viewX - width * 0.1D;
        y = viewY - height * 0.1D;
        regionWidth = width * 1.2D;
        regionHeight = height * 1.2D;
        if (maskElement.Attribute("x") != null) {
            if (!TryViewportLength(maskElement.Attribute("x")!.Value, width, out x, out bool percentage)) return false;
            if (percentage) x += viewX;
        }
        if (maskElement.Attribute("y") != null) {
            if (!TryViewportLength(maskElement.Attribute("y")!.Value, height, out y, out bool percentage)) return false;
            if (percentage) y += viewY;
        }
        if (maskElement.Attribute("width") != null
            && !TryViewportLength(maskElement.Attribute("width")!.Value, width, out regionWidth, out _)) return false;
        if (maskElement.Attribute("height") != null
            && !TryViewportLength(maskElement.Attribute("height")!.Value, height, out regionHeight, out _)) return false;

        x -= viewX;
        y -= viewY;
        return regionWidth > 0D && regionHeight > 0D;
    }

    private static OfficeDrawing ClipMaskToRegion(
        OfficeDrawing content,
        double regionX,
        double regionY,
        double regionWidth,
        double regionHeight,
        OfficeTransform transform) {
        var region = new List<OfficePoint>(4) {
            transform.TransformPoint(new OfficePoint(regionX, regionY)),
            transform.TransformPoint(new OfficePoint(regionX + regionWidth, regionY)),
            transform.TransformPoint(new OfficePoint(regionX + regionWidth, regionY + regionHeight)),
            transform.TransformPoint(new OfficePoint(regionX, regionY + regionHeight))
        };
        region = ClipPolygon(region, MaskClipBoundary.Left, 0D);
        region = ClipPolygon(region, MaskClipBoundary.Top, 0D);
        region = ClipPolygon(region, MaskClipBoundary.Right, content.Width);
        region = ClipPolygon(region, MaskClipBoundary.Bottom, content.Height);
        if (region.Count < 3) return new OfficeDrawing(content.Width, content.Height);

        double left = double.MaxValue;
        double top = double.MaxValue;
        double right = double.MinValue;
        double bottom = double.MinValue;
        var commands = new List<OfficePathCommand>(region.Count + 1);
        for (int index = 0; index < region.Count; index++) {
            OfficePoint point = region[index];
            left = Math.Min(left, point.X);
            top = Math.Min(top, point.Y);
            right = Math.Max(right, point.X);
            bottom = Math.Max(bottom, point.Y);
            commands.Add(index == 0
                ? OfficePathCommand.MoveTo(point)
                : OfficePathCommand.LineTo(point));
        }
        if (right - left <= 0.000001D || bottom - top <= 0.000001D) {
            return new OfficeDrawing(content.Width, content.Height);
        }
        commands.Add(OfficePathCommand.Close());

        var clipped = new OfficeDrawing(content.Width, content.Height);
        clipped.AddClippedDrawing(
            content,
            left,
            top,
            OfficeClipPath.Path(commands),
            -left,
            -top);
        return clipped;
    }

    private static List<OfficePoint> ClipPolygon(
        IReadOnlyList<OfficePoint> source,
        MaskClipBoundary boundary,
        double boundaryValue) {
        var result = new List<OfficePoint>(source.Count + 2);
        if (source.Count == 0) return result;
        OfficePoint previous = source[source.Count - 1];
        bool previousInside = IsInsideMaskBoundary(previous, boundary, boundaryValue);
        for (int index = 0; index < source.Count; index++) {
            OfficePoint current = source[index];
            bool currentInside = IsInsideMaskBoundary(current, boundary, boundaryValue);
            if (currentInside != previousInside) {
                result.Add(IntersectMaskBoundary(previous, current, boundary, boundaryValue));
            }
            if (currentInside) result.Add(current);
            previous = current;
            previousInside = currentInside;
        }
        return result;
    }

    private static bool IsInsideMaskBoundary(OfficePoint point, MaskClipBoundary boundary, double value) {
        switch (boundary) {
            case MaskClipBoundary.Left: return point.X >= value;
            case MaskClipBoundary.Top: return point.Y >= value;
            case MaskClipBoundary.Right: return point.X <= value;
            default: return point.Y <= value;
        }
    }

    private static OfficePoint IntersectMaskBoundary(
        OfficePoint start,
        OfficePoint end,
        MaskClipBoundary boundary,
        double value) {
        if (boundary is MaskClipBoundary.Left or MaskClipBoundary.Right) {
            double ratio = (value - start.X) / (end.X - start.X);
            return new OfficePoint(value, start.Y + ratio * (end.Y - start.Y));
        }
        double verticalRatio = (value - start.Y) / (end.Y - start.Y);
        return new OfficePoint(start.X + verticalRatio * (end.X - start.X), value);
    }

    private enum MaskClipBoundary {
        Left,
        Top,
        Right,
        Bottom
    }

    private static string? ReadPresentationProperty(XElement element, string propertyName) {
        string? value = element.Attribute(propertyName)?.Value;
        string? style = element.Attribute("style")?.Value;
        if (string.IsNullOrWhiteSpace(style)) return value;
        foreach (string declaration in style!.Split(';')) {
            int colon = declaration.IndexOf(':');
            if (colon <= 0 || !declaration.Substring(0, colon).Trim().Equals(propertyName, StringComparison.OrdinalIgnoreCase)) continue;
            value = declaration.Substring(colon + 1).Trim();
        }
        return value;
    }

    private static bool TryParseBlendMode(string value, out OfficeBlendMode mode) {
        switch (value.Trim().ToLowerInvariant()) {
            case "normal": mode = OfficeBlendMode.Normal; return true;
            case "multiply": mode = OfficeBlendMode.Multiply; return true;
            case "screen": mode = OfficeBlendMode.Screen; return true;
            case "overlay": mode = OfficeBlendMode.Overlay; return true;
            case "darken": mode = OfficeBlendMode.Darken; return true;
            case "lighten": mode = OfficeBlendMode.Lighten; return true;
            case "color-dodge": mode = OfficeBlendMode.ColorDodge; return true;
            case "color-burn": mode = OfficeBlendMode.ColorBurn; return true;
            case "hard-light": mode = OfficeBlendMode.HardLight; return true;
            case "soft-light": mode = OfficeBlendMode.SoftLight; return true;
            case "difference": mode = OfficeBlendMode.Difference; return true;
            case "exclusion": mode = OfficeBlendMode.Exclusion; return true;
            case "hue": mode = OfficeBlendMode.Hue; return true;
            case "saturation": mode = OfficeBlendMode.Saturation; return true;
            case "color": mode = OfficeBlendMode.Color; return true;
            case "luminosity": mode = OfficeBlendMode.Luminosity; return true;
            default: mode = OfficeBlendMode.Normal; return false;
        }
    }
}
