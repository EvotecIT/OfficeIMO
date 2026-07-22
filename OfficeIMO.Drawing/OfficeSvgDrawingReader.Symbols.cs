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
        int depth,
        ref int visited,
        ref int pathCommands,
        ref int unsupported) {
        if (!TryParseNumberList(symbol.Attribute("viewBox")?.Value, out IReadOnlyList<double> viewBox)
            || viewBox.Count != 4
            || viewBox[2] <= 0D
            || viewBox[3] <= 0D
            || !IsSupportedSvgViewport(viewBox[2], viewBox[3])
            || !TrySymbolLength(use, symbol, "width", viewBox[2], out double width)
            || !TrySymbolLength(use, symbol, "height", viewBox[3], out double height)
            || !TryOptionalUseLength(use, "x", out double x)
            || !TryOptionalUseLength(use, "y", out double y)
            || width <= 0D
            || height <= 0D
            || !IsSupportedSvgViewport(width, height)) {
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

        var scene = new OfficeDrawing(viewBox[2], viewBox[3]);
        SvgPaintContext style = ResolvePaintContext(symbol, inheritedStyle, paintServers, ref unsupported);
        OfficeTransform symbolTransform = ResolveTransform(symbol, OfficeTransform.Identity, viewBox[0], viewBox[1], ref unsupported);
        AddChildren(symbol, scene, style, paintServers, references, symbolTransform, viewBox[0], viewBox[1],
            maximumElements, depth, ref visited, ref pathCommands, ref unsupported);

        OfficeTransform viewportTransform;
        if (alignment == SvgAspectAlignment.None) {
            viewportTransform = OfficeTransform.Scale(width / viewBox[2], height / viewBox[3]);
        } else {
            double scale = slice
                ? Math.Max(width / viewBox[2], height / viewBox[3])
                : Math.Min(width / viewBox[2], height / viewBox[3]);
            double remainingX = width - (viewBox[2] * scale);
            double remainingY = height - (viewBox[3] * scale);
            ResolveAlignmentFactors(alignment, out double alignX, out double alignY);
            viewportTransform = OfficeTransform.Scale(scale, scale)
                .Then(OfficeTransform.Translate(remainingX * alignX, remainingY * alignY));
        }

        var viewport = new OfficeDrawing(width, height);
        viewport.AddEffectDrawing(scene, viewportTransform);
        var clipped = new OfficeDrawing(width, height);
        clipped.AddClippedDrawing(viewport, 0D, 0D, OfficeClipPath.Rectangle(width, height));
        drawing.AddEffectDrawing(clipped, OfficeTransform.Translate(x, y).Then(inheritedTransform));
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

        string[] parts = value!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        int index = parts.Length > 0 && parts[0].Equals("defer", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
        if (index >= parts.Length || !TryParseAspectAlignment(parts[index++], out alignment)) return false;
        if (alignment == SvgAspectAlignment.None) return index == parts.Length;
        if (index == parts.Length) return true;
        if (index + 1 != parts.Length) return false;
        if (parts[index].Equals("meet", StringComparison.OrdinalIgnoreCase)) return true;
        if (!parts[index].Equals("slice", StringComparison.OrdinalIgnoreCase)) return false;
        slice = true;
        return true;
    }

    private static bool TryParseAspectAlignment(string value, out SvgAspectAlignment alignment) {
        switch (value.ToLowerInvariant()) {
            case "none": alignment = SvgAspectAlignment.None; return true;
            case "xminymin": alignment = SvgAspectAlignment.XMinYMin; return true;
            case "xmidymin": alignment = SvgAspectAlignment.XMidYMin; return true;
            case "xmaxymin": alignment = SvgAspectAlignment.XMaxYMin; return true;
            case "xminymid": alignment = SvgAspectAlignment.XMinYMid; return true;
            case "xmidymid": alignment = SvgAspectAlignment.XMidYMid; return true;
            case "xmaxymid": alignment = SvgAspectAlignment.XMaxYMid; return true;
            case "xminymax": alignment = SvgAspectAlignment.XMinYMax; return true;
            case "xmidymax": alignment = SvgAspectAlignment.XMidYMax; return true;
            case "xmaxymax": alignment = SvgAspectAlignment.XMaxYMax; return true;
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
