using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static OfficeDrawing FitSvgViewport(OfficeDrawing scene, double width, double height,
        OfficeTransform transform, double maximumDimension, double maximumPixels, ref int unsupported) {
        if (HasOverflowingLocalShapes(scene) && transform.TryInvert(out OfficeTransform inverse)) {
            var visible = inverse.TransformRectangleBounds(0D, 0D, width, height);
            if (scene.TryExpandViewportCanvas(visible.Left, visible.Top, visible.Right, visible.Bottom,
                    maximumDimension, maximumPixels, out OfficeDrawing expanded, out double left, out double top)) {
                scene = expanded;
                transform = OfficeTransform.Translate(left, top).Then(transform);
            } else {
                unsupported++;
            }
        }
        return new OfficeDrawing(width, height).AddEffectDrawing(scene, transform);
    }

    // Only newly retained, out-of-canvas geometry needs an additional root wrapper. Keep the
    // existing flat scene contract for ordinary primitives, including their stroke metadata.
    private static bool HasNewlyRetainedSvgGeometry(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingShape shape &&
                (shape.X < 0D || shape.Y < 0D || shape.X + shape.Shape.Width > drawing.Width ||
                 shape.Y + shape.Shape.Height > drawing.Height)) return true;
            if (element is OfficeDrawingEffectGroup effect && HasNewlyRetainedSvgGeometry(effect.InnerDrawing)) return true;
            if (element is OfficeDrawingGroup group && HasNewlyRetainedSvgGeometry(group.InnerDrawing)) return true;
        }
        return false;
    }

    // Viewport fitting must also account for transformed paint and stroke extent.
    private static bool HasOverflowingLocalShapes(OfficeDrawing drawing) {
        return HasOverflowingLocalShapes(drawing, OfficeTransform.Identity, drawing.Width, drawing.Height);
    }

    private static bool HasOverflowingLocalShapes(OfficeDrawing drawing, OfficeTransform parent, double width, double height) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingShape shape) {
                double stroke = shape.Shape.StrokeColor.HasValue || shape.Shape.StrokeGradient != null || shape.Shape.StrokeRadialGradient != null
                    ? shape.Shape.StrokeWidth / 2D : 0D;
                if (shape.Shape.StrokeLineJoin == null || shape.Shape.StrokeLineJoin == OfficeStrokeLineJoin.Miter)
                    stroke *= Math.Max(1D, shape.Shape.StrokeMiterLimit);
                OfficeTransform transform = (shape.Shape.Transform ?? OfficeTransform.Identity)
                    .Then(OfficeTransform.Translate(shape.X, shape.Y)).Then(parent);
                var bounds = transform.TransformRectangleBounds(-stroke, -stroke,
                    shape.Shape.Width + stroke * 2D, shape.Shape.Height + stroke * 2D);
                if (bounds.Left < 0D || bounds.Top < 0D || bounds.Right > width || bounds.Bottom > height) return true;
            }
            if (element is OfficeDrawingEffectGroup effect && HasOverflowingLocalShapes(effect.InnerDrawing,
                    effect.Transform.Then(parent), width, height)) return true;
            if (element is OfficeDrawingGroup group) {
                OfficeTransform transform = OfficeTransform.Translate(group.X + group.ContentOffsetX, group.Y + group.ContentOffsetY);
                if (group.FrameTransform.HasValue) transform = transform.Then(group.FrameTransform.Value.CreateDestinationTransform());
                if (HasOverflowingLocalShapes(group.InnerDrawing, transform.Then(parent), width, height)) return true;
            }
        }
        return false;
    }

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

        string? clipValue = ReadPresentationProperty(element, "clip-path");
        if (!string.IsNullOrWhiteSpace(clipValue) && !clipValue!.Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) {
            if (TryApplySvgClip(scene, clipValue!, references, paintServers, childTransform, childViewX, childViewY,
                    maximumElements, ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported,
                    out OfficeDrawing? clippedScene)) scene = clippedScene!;
            else unsupported++;
        }

        OfficeTransform viewportTransform = ResolveViewportTransform(
            childViewWidth, childViewHeight, width, height, alignment, slice);
        OfficeDrawing viewport = FitSvgViewport(scene, width, height, viewportTransform,
            maximumViewportDimension, maximumViewportPixels, ref unsupported);
        viewport.Fonts.AddRange(drawing.Fonts);
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
