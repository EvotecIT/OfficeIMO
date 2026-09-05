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
        ref bool pathCommandLimitExceeded,
        ref int unsupported,
        out OfficeBlendMode blendMode,
        out OfficeDrawingSoftMask? softMask,
        out SvgFilterEffect? filterEffect) {
        blendMode = OfficeBlendMode.Normal;
        softMask = null;
        filterEffect = null;
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

        string? filterValue = ReadPresentationProperty(element, "filter");
        if (!string.IsNullOrWhiteSpace(filterValue)
            && !filterValue!.Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) {
            if (TryResolveSvgFilter(filterValue, references, out filterEffect)) hasEffects = true;
            else unsupported++;
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
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
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

    private static bool TryResolveSvgFilter(
        string filterValue,
        SvgElementReferenceRegistry references,
        out SvgFilterEffect? effect) {
        effect = null;
        if (!references.TryEnterLocal(filterValue, out string filterId, out XElement? filterElement)) return false;
        try {
            if (filterElement == null
                || !filterElement.Name.LocalName.Equals("filter", StringComparison.OrdinalIgnoreCase)) return false;
            List<XElement> primitives = filterElement.Elements().ToList();
            if (primitives.Count == 1
                && primitives[0].Name.LocalName.Equals("feDropShadow", StringComparison.OrdinalIgnoreCase)) {
                return TryParseDropShadowFilter(primitives[0], out effect);
            }
            if (primitives.Count == 0 || primitives.Count > 2) return false;
            if (primitives.Any(HasUnsupportedFilterRouting)) return false;

            double blur = 0D;
            double offsetX = 0D;
            double offsetY = 0D;
            bool hasBlur = false;
            bool hasOffset = false;
            foreach (XElement primitive in primitives) {
                string name = primitive.Name.LocalName;
                if (name.Equals("feGaussianBlur", StringComparison.OrdinalIgnoreCase)) {
                    if (hasBlur || hasOffset || !TryParseStdDeviation(primitive.Attribute("stdDeviation")?.Value, out blur)) return false;
                    hasBlur = true;
                } else if (name.Equals("feOffset", StringComparison.OrdinalIgnoreCase)) {
                    if (hasOffset
                        || !TryParseFiniteNumber(primitive.Attribute("dx")?.Value, 0D, out offsetX)
                        || !TryParseFiniteNumber(primitive.Attribute("dy")?.Value, 0D, out offsetY)) return false;
                    hasOffset = true;
                } else {
                    return false;
                }
            }
            if (!hasBlur && !hasOffset) return false;
            effect = new SvgFilterEffect(
                hasBlur ? SvgFilterEffectKind.GaussianBlur : SvgFilterEffectKind.Offset,
                offsetX,
                offsetY,
                blur,
                OfficeColor.Black,
                1D);
            return true;
        } finally {
            references.Exit(filterId);
        }
    }

    private static bool TryParseDropShadowFilter(XElement primitive, out SvgFilterEffect? effect) {
        effect = null;
        string? standardDeviation = primitive.Attribute("stdDeviation")?.Value;
        if (HasUnsupportedFilterRouting(primitive)
            || !TryParseFiniteNumber(primitive.Attribute("dx")?.Value, 2D, out double offsetX)
            || !TryParseFiniteNumber(primitive.Attribute("dy")?.Value, 2D, out double offsetY)
            || !TryParseStdDeviation(standardDeviation, out double blur)) return false;
        if (string.IsNullOrWhiteSpace(standardDeviation)) blur = 2D;

        string colorValue = ReadPresentationProperty(primitive, "flood-color") ?? "black";
        if (!OfficeColor.TryParseCss(colorValue, out OfficeColor color)) return false;
        if (!TryParseFiniteNumber(ReadPresentationProperty(primitive, "flood-opacity"), 1D, out double opacity)
            || opacity < 0D || opacity > 1D) return false;
        opacity *= color.A / 255D;
        effect = new SvgFilterEffect(
            SvgFilterEffectKind.DropShadow,
            offsetX,
            offsetY,
            blur,
            OfficeColor.FromRgb(color.R, color.G, color.B),
            opacity);
        return true;
    }

    private static bool HasUnsupportedFilterRouting(XElement primitive) =>
        primitive.Attribute("in") != null
        || primitive.Attribute("in2") != null
        || primitive.Attribute("result") != null;

    private static bool TryParseStdDeviation(string? value, out double deviation) {
        deviation = 0D;
        if (string.IsNullOrWhiteSpace(value)) return true;
        if (!TryParseNumberList(value, out IReadOnlyList<double> values)
            || values.Count is < 1 or > 2
            || values.Any(number => number < 0D || double.IsNaN(number) || double.IsInfinity(number))) return false;
        deviation = values.Max();
        return true;
    }

    private static bool TryParseFiniteNumber(string? value, double fallback, out double result) {
        result = fallback;
        if (string.IsNullOrWhiteSpace(value)) return true;
        return double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out result)
            && !double.IsNaN(result)
            && !double.IsInfinity(result);
    }

    private static bool TryApplySvgFilter(
        OfficeDrawing source,
        SvgFilterEffect? effect,
        OfficeTransform transform,
        int maximumElements,
        ref int visited,
        ref int unsupported,
        out OfficeDrawing result) {
        result = source;
        if (effect == null) return true;
        if (effect.Kind == SvgFilterEffectKind.DropShadow && ContainsUntintableFilterPaint(source)) {
            unsupported++;
            return false;
        }

        IReadOnlyList<OfficePoint> samples = CreateSvgFilterSamples(effect.BlurRadius);
        int sourceElements = CountDrawingElements(source, maximumElements);
        int copyCount = effect.Kind == SvgFilterEffectKind.DropShadow ? samples.Count
            : effect.Kind == SvgFilterEffectKind.GaussianBlur ? samples.Count - 1
            : 0;
        long additionalElements = (long)sourceElements * copyCount;
        if (additionalElements > maximumElements - visited) {
            unsupported++;
            return false;
        }
        visited += (int)additionalElements;

        var filtered = new OfficeDrawing(source.Width, source.Height);
        filtered.Fonts.AddRange(source.Fonts);
        if (effect.Kind == SvgFilterEffectKind.Offset) {
            OfficePoint offset = TransformSvgEffectVector(transform, effect.OffsetX, effect.OffsetY);
            filtered.AddEffectDrawing(source, OfficeTransform.Translate(offset.X, offset.Y));
        } else {
            OfficeDrawing paint = source.Clone();
            if (effect.Kind == SvgFilterEffectKind.DropShadow) paint.ApplyColorTint(effect.Color);
            var blurred = new OfficeDrawing(source.Width, source.Height);
            blurred.Fonts.AddRange(source.Fonts);
            for (int sampleIndex = 0; sampleIndex < samples.Count; sampleIndex++) {
                OfficePoint sample = samples[sampleIndex];
                OfficePoint offset = TransformSvgEffectVector(
                    transform,
                    effect.OffsetX + sample.X,
                    effect.OffsetY + sample.Y);
                blurred.AddEffectDrawing(
                    paint,
                    OfficeTransform.Translate(offset.X, offset.Y),
                    ResolveSvgFilterSampleOpacity(effect.Opacity, samples.Count, sampleIndex));
            }
            filtered.AddEffectDrawing(blurred, OfficeTransform.Identity);
            if (effect.Kind == SvgFilterEffectKind.DropShadow) {
                filtered.AddEffectDrawing(source, OfficeTransform.Identity);
            }
        }
        result = filtered;
        return true;
    }

    private static OfficePoint TransformSvgEffectVector(OfficeTransform transform, double x, double y) =>
        new OfficePoint(
            transform.M11 * x + transform.M21 * y,
            transform.M12 * x + transform.M22 * y);

    private static IReadOnlyList<OfficePoint> CreateSvgFilterSamples(double blurRadius) {
        if (blurRadius <= 0.0001D) return new[] { new OfficePoint(0D, 0D) };
        double radius = blurRadius * 1.3D;
        double diagonal = radius * 0.7071067811865476D;
        return new[] {
            new OfficePoint(0D, 0D),
            new OfficePoint(radius, 0D),
            new OfficePoint(-radius, 0D),
            new OfficePoint(0D, radius),
            new OfficePoint(0D, -radius),
            new OfficePoint(diagonal, diagonal),
            new OfficePoint(diagonal, -diagonal),
            new OfficePoint(-diagonal, diagonal),
            new OfficePoint(-diagonal, -diagonal)
        };
    }

    private static int CountDrawingElements(OfficeDrawing drawing, int limit) {
        int count = 0;
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (++count >= limit) return limit;
            OfficeDrawing? child = element is OfficeDrawingGroup group ? group.InnerDrawing
                : element is OfficeDrawingEffectGroup effectGroup ? effectGroup.InnerDrawing
                : null;
            if (child == null) continue;
            count += CountDrawingElements(child, limit - count);
            if (count >= limit) return limit;
        }
        return count;
    }

    private static bool ContainsUntintableFilterPaint(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingImage or OfficeDrawingImagePattern) return true;
            OfficeDrawing? child = element is OfficeDrawingGroup group ? group.InnerDrawing
                : element is OfficeDrawingEffectGroup effectGroup ? effectGroup.InnerDrawing
                : null;
            if (child != null && ContainsUntintableFilterPaint(child)) return true;
        }
        return false;
    }

    private static double ResolveSvgFilterSampleOpacity(double opacity, int sampleCount, int sampleIndex) {
        if (sampleCount <= 1) return opacity;
        return Math.Max(0D, Math.Min(1D, opacity * (sampleIndex == 0 ? 0.7D : 0.18D)));
    }

    private enum SvgFilterEffectKind {
        DropShadow,
        GaussianBlur,
        Offset
    }

    private sealed class SvgFilterEffect {
        internal SvgFilterEffect(
            SvgFilterEffectKind kind,
            double offsetX,
            double offsetY,
            double blurRadius,
            OfficeColor color,
            double opacity) {
            Kind = kind;
            OffsetX = offsetX;
            OffsetY = offsetY;
            BlurRadius = blurRadius;
            Color = color;
            Opacity = opacity;
        }

        internal SvgFilterEffectKind Kind { get; }
        internal double OffsetX { get; }
        internal double OffsetY { get; }
        internal double BlurRadius { get; }
        internal OfficeColor Color { get; }
        internal double Opacity { get; }
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
        int selectedPriority = -1;
        foreach (string declaration in style!.Split(';')) {
            int colon = declaration.IndexOf(':');
            if (colon <= 0 || !declaration.Substring(0, colon).Trim().Equals(propertyName, StringComparison.OrdinalIgnoreCase)) continue;
            string candidate = NormalizeInlineStyleValue(declaration.Substring(colon + 1), out int priority);
            if (priority < selectedPriority) continue;
            value = candidate;
            selectedPriority = priority;
        }
        return value;
    }

    private static string NormalizeInlineStyleValue(string value, out int priority) {
        string normalized = value.Trim();
        priority = 0;
        int priorityStart = normalized.LastIndexOf('!');
        if (priorityStart < 0) return normalized;
        if (!IsCssImportantPriority(normalized, priorityStart + 1)) {
            if (HasPotentialSvgUrlFunction(normalized)) priority = 2;
            return normalized;
        }
        priority = 1;
        return normalized.Substring(0, priorityStart).TrimEnd();
    }

    private static bool IsCssImportantPriority(string value, int start) {
        int index = start;
        SkipCssWhitespaceAndComments(value, ref index);
        const string important = "important";
        if (index + important.Length > value.Length
            || !value.Substring(index, important.Length).Equals(important, StringComparison.OrdinalIgnoreCase)) return false;
        index += important.Length;
        SkipCssWhitespaceAndComments(value, ref index);
        return index == value.Length;
    }

    private static void SkipCssWhitespaceAndComments(string value, ref int index) {
        while (index < value.Length) {
            if (char.IsWhiteSpace(value[index])) {
                index++;
                continue;
            }
            if (index + 1 >= value.Length || value[index] != '/' || value[index + 1] != '*') return;
            int commentEnd = value.IndexOf("*/", index + 2, StringComparison.Ordinal);
            if (commentEnd < 0) {
                index = value.Length;
                return;
            }
            index = commentEnd + 2;
        }
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
