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
        ref int visited,
        ref int unsupported) {
        if (!TryParseNumberList(symbol.Attribute("viewBox")?.Value, out IReadOnlyList<double> viewBox)
            || viewBox.Count != 4
            || viewBox[2] <= 0D
            || viewBox[3] <= 0D
            || !TrySymbolLength(use, symbol, "width", viewBox[2], out double width)
            || !TrySymbolLength(use, symbol, "height", viewBox[3], out double height)
            || !TryOptionalUseLength(use, "x", out double x)
            || !TryOptionalUseLength(use, "y", out double y)
            || width <= 0D
            || height <= 0D) {
            unsupported++;
            return;
        }

        string aspect = (use.Attribute("preserveAspectRatio")?.Value
            ?? symbol.Attribute("preserveAspectRatio")?.Value
            ?? "xMidYMid meet").Trim();
        bool stretch = aspect.Equals("none", StringComparison.OrdinalIgnoreCase);
        if (!stretch && !aspect.Equals("xMidYMid meet", StringComparison.OrdinalIgnoreCase) && !aspect.Equals("xMidYMid", StringComparison.OrdinalIgnoreCase)) {
            unsupported++;
            return;
        }

        var scene = new OfficeDrawing(viewBox[2], viewBox[3]);
        SvgPaintContext style = ResolvePaintContext(symbol, inheritedStyle, paintServers, ref unsupported);
        OfficeTransform symbolTransform = ResolveTransform(symbol, OfficeTransform.Identity, viewBox[0], viewBox[1], ref unsupported);
        AddChildren(symbol, scene, style, paintServers, references, symbolTransform, viewBox[0], viewBox[1], ref visited, ref unsupported);

        OfficeTransform placement;
        if (stretch) {
            placement = OfficeTransform.Scale(width / viewBox[2], height / viewBox[3])
                .Then(OfficeTransform.Translate(x, y));
        } else {
            double scale = Math.Min(width / viewBox[2], height / viewBox[3]);
            double offsetX = x + ((width - (viewBox[2] * scale)) / 2D);
            double offsetY = y + ((height - (viewBox[3] * scale)) / 2D);
            placement = OfficeTransform.Scale(scale, scale).Then(OfficeTransform.Translate(offsetX, offsetY));
        }
        drawing.AddEffectDrawing(scene, placement.Then(inheritedTransform));
    }

    private static bool TrySymbolLength(XElement use, XElement symbol, string name, double fallback, out double value) {
        string? text = use.Attribute(name)?.Value ?? symbol.Attribute(name)?.Value;
        if (string.IsNullOrWhiteSpace(text)) {
            value = fallback;
            return true;
        }
        return TrySvgLength(text, out value);
    }
}
