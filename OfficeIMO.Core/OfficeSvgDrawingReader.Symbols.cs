using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static void AddReferencedSymbol(
        XElement use,
        XElement symbol,
        OfficeDrawing drawing,
        SvgPaintContext inheritedStyle,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported) {
        if (!TryParseNumberList(symbol.Attribute("viewBox")?.Value, out IReadOnlyList<double> viewBox)
            || viewBox.Count != 4
            || viewBox[2] <= 0D
            || viewBox[3] <= 0D
            || !IsSupportedSvgViewport(viewBox[2], viewBox[3], maximumViewportDimension, maximumViewportPixels)
            || !TrySymbolLength(use, symbol, "width", viewBox[2], out double width)
            || !TrySymbolLength(use, symbol, "height", viewBox[3], out double height)
            || !TryOptionalUseLength(use, "x", out double x)
            || !TryOptionalUseLength(use, "y", out double y)
            || width <= 0D
            || height <= 0D
            || !IsSupportedSvgViewport(width, height, maximumViewportDimension, maximumViewportPixels)) {
            unsupported++;
            return;
        }

        if (!TryParsePreserveAspectRatio(
                use.Attribute("preserveAspectRatio")?.Value ?? symbol.Attribute("preserveAspectRatio")?.Value,
                out SvgAspectAlignment alignment,
                out bool slice)) {
            unsupported++;
            return;
        }

        if (!references.TryChargeNestedViewport(width, height, viewBox[2], viewBox[3])) {
            unsupported++;
            return;
        }

        var scene = new OfficeDrawing(viewBox[2], viewBox[3]);
        scene.Fonts.AddRange(drawing.Fonts);
        SvgPaintContext style = ResolvePaintContext(symbol, inheritedStyle, paintServers, ref unsupported);
        OfficeTransform symbolTransform = ResolveTransform(symbol, OfficeTransform.Identity, viewBox[0], viewBox[1], ref unsupported);
        AddChildren(symbol, scene, style, paintServers, references, symbolTransform, viewBox[0], viewBox[1],
            maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
            ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);

        OfficeTransform viewportTransform = ResolveViewportTransform(viewBox[2], viewBox[3], width, height, alignment, slice);

        OfficeDrawing viewport = FitSvgViewport(scene, width, height, viewportTransform,
            maximumViewportDimension, maximumViewportPixels, ref unsupported, out double retainedScenePixels);
        if (!references.TryChargeNestedViewportExpansion(retainedScenePixels - viewBox[2] * viewBox[3])) {
            unsupported++;
            return;
        }
        var clipped = new OfficeDrawing(width, height);
        clipped.AddClippedDrawing(viewport, 0D, 0D, OfficeClipPath.Rectangle(width, height));
        drawing.AddEffectDrawing(clipped, OfficeTransform.Translate(x, y).Then(inheritedTransform));
    }

    private static OfficeTransform ResolveViewportTransform(
        double viewWidth,
        double viewHeight,
        double viewportWidth,
        double viewportHeight,
        SvgAspectAlignment alignment,
        bool slice) {
        if (alignment == SvgAspectAlignment.None) {
            return OfficeTransform.Scale(viewportWidth / viewWidth, viewportHeight / viewHeight);
        }

        double scale = slice
            ? Math.Max(viewportWidth / viewWidth, viewportHeight / viewHeight)
            : Math.Min(viewportWidth / viewWidth, viewportHeight / viewHeight);
        double remainingX = viewportWidth - (viewWidth * scale);
        double remainingY = viewportHeight - (viewHeight * scale);
        ResolveAlignmentFactors(alignment, out double alignX, out double alignY);
        return OfficeTransform.Scale(scale, scale)
            .Then(OfficeTransform.Translate(remainingX * alignX, remainingY * alignY));
    }

    private static bool TrySymbolLength(XElement use, XElement symbol, string name, double fallback, out double value) {
        string? text = use.Attribute(name)?.Value ?? symbol.Attribute(name)?.Value;
        if (string.IsNullOrWhiteSpace(text)) {
            value = fallback;
            return true;
        }
        return TrySvgLength(text, out value);
    }

    private static bool TryParsePreserveAspectRatio(string? value, out SvgAspectAlignment alignment, out bool slice) {
        alignment = SvgAspectAlignment.XMidYMid;
        slice = false;
        if (string.IsNullOrWhiteSpace(value)) return true;

        if (ContainsNonSvgCssWhitespace(value!)) return false;
        string[] parts = value!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        int index = parts.Length > 0 && parts[0].Equals("defer", StringComparison.Ordinal) ? 1 : 0;
        if (index >= parts.Length || !TryParseAspectAlignment(parts[index++], out alignment)) return false;
        if (alignment == SvgAspectAlignment.None) {
            if (index == parts.Length) return true;
            return index + 1 == parts.Length &&
                (parts[index].Equals("meet", StringComparison.Ordinal) ||
                 parts[index].Equals("slice", StringComparison.Ordinal));
        }
        if (index == parts.Length) return true;
        if (index + 1 != parts.Length) return false;
        if (parts[index].Equals("meet", StringComparison.Ordinal)) return true;
        if (!parts[index].Equals("slice", StringComparison.Ordinal)) return false;
        slice = true;
        return true;
    }

    private static bool TryParseAspectAlignment(string value, out SvgAspectAlignment alignment) {
        switch (value) {
            case "none": alignment = SvgAspectAlignment.None; return true;
            case "xMinYMin": alignment = SvgAspectAlignment.XMinYMin; return true;
            case "xMidYMin": alignment = SvgAspectAlignment.XMidYMin; return true;
            case "xMaxYMin": alignment = SvgAspectAlignment.XMaxYMin; return true;
            case "xMinYMid": alignment = SvgAspectAlignment.XMinYMid; return true;
            case "xMidYMid": alignment = SvgAspectAlignment.XMidYMid; return true;
            case "xMaxYMid": alignment = SvgAspectAlignment.XMaxYMid; return true;
            case "xMinYMax": alignment = SvgAspectAlignment.XMinYMax; return true;
            case "xMidYMax": alignment = SvgAspectAlignment.XMidYMax; return true;
            case "xMaxYMax": alignment = SvgAspectAlignment.XMaxYMax; return true;
            default:
                alignment = default;
                return false;
        }
    }

    private static void ResolveAlignmentFactors(SvgAspectAlignment alignment, out double x, out double y) {
        x = alignment is SvgAspectAlignment.XMinYMin or SvgAspectAlignment.XMinYMid or SvgAspectAlignment.XMinYMax ? 0D
            : alignment is SvgAspectAlignment.XMaxYMin or SvgAspectAlignment.XMaxYMid or SvgAspectAlignment.XMaxYMax ? 1D
            : 0.5D;
        y = alignment is SvgAspectAlignment.XMinYMin or SvgAspectAlignment.XMidYMin or SvgAspectAlignment.XMaxYMin ? 0D
            : alignment is SvgAspectAlignment.XMinYMax or SvgAspectAlignment.XMidYMax or SvgAspectAlignment.XMaxYMax ? 1D
            : 0.5D;
    }

    private enum SvgAspectAlignment {
        None,
        XMinYMin,
        XMidYMin,
        XMaxYMin,
        XMinYMid,
        XMidYMid,
        XMaxYMid,
        XMinYMax,
        XMidYMax,
        XMaxYMax
    }
}
