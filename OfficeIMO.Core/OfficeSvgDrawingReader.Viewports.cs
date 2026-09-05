using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryAddNestedSvgViewport(
        XElement element,
        OfficeDrawing drawing,
        SvgPaintContext style,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform transform,
        double parentViewX,
        double parentViewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported) {
        double x = ReadViewportCoordinate(element, "x", parentViewX, drawing.Width);
        double y = ReadViewportCoordinate(element, "y", parentViewY, drawing.Height);
        if (!TryNestedViewportLength(element.Attribute("width")?.Value, drawing.Width, out double width)
            || !TryNestedViewportLength(element.Attribute("height")?.Value, drawing.Height, out double height)
            || width <= 0D
            || height <= 0D
            || !IsSupportedSvgViewport(width, height, maximumViewportDimension, maximumViewportPixels)) {
            return false;
        }

        double childViewX = 0D;
        double childViewY = 0D;
        double childViewWidth = width;
        double childViewHeight = height;
        string? viewBoxText = element.Attribute("viewBox")?.Value;
        if (!string.IsNullOrWhiteSpace(viewBoxText)) {
            if (!TryParseNumberList(viewBoxText, out IReadOnlyList<double> viewBox)
                || viewBox.Count != 4
                || viewBox[2] <= 0D
                || viewBox[3] <= 0D
                || !IsSupportedSvgViewport(viewBox[2], viewBox[3], maximumViewportDimension, maximumViewportPixels)) {
                return false;
            }
            childViewX = viewBox[0];
            childViewY = viewBox[1];
            childViewWidth = viewBox[2];
            childViewHeight = viewBox[3];
        }
        if (!TryParsePreserveAspectRatio(element.Attribute("preserveAspectRatio")?.Value,
                out SvgAspectAlignment alignment, out bool slice)) return false;

        bool hasEffects = TryResolveSvgEffects(
            element,
            width,
            height,
            style,
            paintServers,
            references,
            transform,
            parentViewX,
            parentViewY,
            maximumElements,
            maximumViewportDimension,
            maximumViewportPixels,
            depth,
            ref visited,
            ref pathCommands,
            ref pathCommandLimitExceeded,
            ref unsupported,
            out OfficeBlendMode blendMode,
            out OfficeDrawingSoftMask? softMask,
            out SvgFilterEffect? filterEffect);

        var scene = new OfficeDrawing(childViewWidth, childViewHeight);
        scene.Fonts.AddRange(drawing.Fonts);
        OfficeTransform childTransform = ResolveTransform(
            element,
            OfficeTransform.Identity,
            childViewX,
            childViewY,
            ref unsupported);
        // The viewport element's own transform is applied to the viewport below. Do not apply it
        // a second time to its local child coordinate system.
        if (element.Attribute("transform") != null) childTransform = OfficeTransform.Identity;
        style.DashPercentageReference = NormalizedSvgDiagonal(childViewWidth, childViewHeight);
        AddChildren(
            element, scene, style, paintServers, references, childTransform, childViewX, childViewY,
            maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
            ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);

        OfficeTransform viewportTransform = ResolveViewportTransform(
            childViewWidth, childViewHeight, width, height, alignment, slice);
        var viewport = new OfficeDrawing(width, height);
        viewport.Fonts.AddRange(drawing.Fonts);
        viewport.AddEffectDrawing(scene, viewportTransform);
        var clipped = new OfficeDrawing(width, height);
        clipped.Fonts.AddRange(drawing.Fonts);
        clipped.AddClippedDrawing(viewport, 0D, 0D, OfficeClipPath.Rectangle(width, height));
        OfficeDrawing content = clipped;
        if (hasEffects) {
            TryApplySvgFilter(content, filterEffect, OfficeTransform.Identity, maximumElements,
                ref visited, ref unsupported, out content);
        }
        drawing.AddEffectDrawing(
            content,
            OfficeTransform.Translate(x, y).Then(transform),
            blendMode,
            softMask);
        return true;
    }

    private static bool TryNestedViewportLength(string? text, double reference, out double value) {
        if (string.IsNullOrWhiteSpace(text)) {
            value = reference;
            return true;
        }
        return TryViewportLength(text, reference, out value, out _);
    }
}
