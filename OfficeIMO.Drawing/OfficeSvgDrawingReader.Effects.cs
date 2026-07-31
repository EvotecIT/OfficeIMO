using System;
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
            hasEffects = true;
            if (!TryParseBlendMode(blendValue!, out blendMode)) unsupported++;
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
            SvgPaintContext maskStyle = ResolvePaintContext(maskElement, SvgPaintContext.Default, paintServers, ref unsupported);
            OfficeTransform maskTransform = ResolveTransform(maskElement, transform, viewX, viewY, ref unsupported);
            AddChildren(maskElement, maskContent, maskStyle, paintServers, references, maskTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                ref visited, ref pathCommands, ref unsupported);
            OfficeDrawing maskDrawing = ClipMaskToRegion(maskContent, regionX, regionY, regionWidth, regionHeight);
            OfficeSoftMaskMode mode = ReadPresentationProperty(maskElement, "mask-type")?.Trim().Equals("alpha", StringComparison.OrdinalIgnoreCase) == true
                ? OfficeSoftMaskMode.Alpha
                : OfficeSoftMaskMode.Luminosity;
            softMask = new OfficeDrawingSoftMask(maskDrawing, mode);
            return true;
        } finally {
            references.Exit(maskId);
        }
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
        x = -width * 0.1D;
        y = -height * 0.1D;
        regionWidth = width * 1.2D;
        regionHeight = height * 1.2D;
        if (maskElement.Attribute("x") != null
            && !TryViewportLength(maskElement.Attribute("x")!.Value, width, out x, out _)) return false;
        if (maskElement.Attribute("y") != null
            && !TryViewportLength(maskElement.Attribute("y")!.Value, height, out y, out _)) return false;
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
        double regionHeight) {
        double left = Math.Max(0D, regionX);
        double top = Math.Max(0D, regionY);
        double right = Math.Min(content.Width, regionX + regionWidth);
        double bottom = Math.Min(content.Height, regionY + regionHeight);
        if (right <= left || bottom <= top) return new OfficeDrawing(content.Width, content.Height);
        if (left <= 0D && top <= 0D && right >= content.Width && bottom >= content.Height) return content;

        var clipped = new OfficeDrawing(content.Width, content.Height);
        clipped.AddClippedDrawing(
            content,
            left,
            top,
            OfficeClipPath.Rectangle(right - left, bottom - top),
            -left,
            -top);
        return clipped;
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
