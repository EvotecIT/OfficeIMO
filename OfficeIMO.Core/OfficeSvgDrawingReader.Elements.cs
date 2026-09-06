using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reads a bounded subset of SVG into the shared dependency-free drawing scene.
/// </summary>
public static partial class OfficeSvgDrawingReader {
    private static void AddChildren(
        XElement parent,
        OfficeDrawing drawing,
        SvgPaintContext inherited,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported) {
        if (depth > MaximumSvgNestingDepth) {
            unsupported++;
            return;
        }
        foreach (XElement element in parent.Elements()) {
            AddElement(element, drawing, inherited, paintServers, references, inheritedTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            if (visited > maximumElements) return;
        }
    }

    private static void AddElement(
        XElement element,
        OfficeDrawing drawing,
        SvgPaintContext inherited,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
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
        bool suppressElementClip = false) {
        visited++;
        if (visited > maximumElements) return;
        string name = element.Name.LocalName.ToLowerInvariant();
        if (name is "title" or "desc" or "metadata" or "style" or "lineargradient" or "radialgradient" or "pattern" or "stop") return;
        if (name == "defs") return;

        string? clipValue = ReadPresentationProperty(element, "clip-path");
        if (name != "svg" && !suppressElementClip && !string.IsNullOrWhiteSpace(clipValue) &&
            !clipValue!.Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) {
            var content = new OfficeDrawing(drawing.Width, drawing.Height);
            content.Fonts.AddRange(drawing.Fonts);
            AddElement(element, content, inherited, paintServers, references, inheritedTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported, suppressElementClip: true);
            OfficeTransform clipTransform = ResolveTransform(element, inheritedTransform, viewX, viewY, ref unsupported);
            if (name == "use" && TryOptionalUseLength(element, "x", out double clipX) && TryOptionalUseLength(element, "y", out double clipY))
                clipTransform = OfficeTransform.Translate(clipX, clipY).Then(clipTransform);
            if (!TryApplySvgClip(content, clipValue!, references, paintServers, clipTransform, viewX, viewY,
                    maximumElements, ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported, out OfficeDrawing? clipped)) {
                unsupported++;
            }
            drawing.AddDrawingForClippedRendering(clipped ?? content, 0D, 0D, null);
            return;
        }

        inherited.DashPercentageReference = NormalizedSvgDiagonal(drawing.Width, drawing.Height);
        SvgPaintContext style = ResolvePaintContext(element, inherited, paintServers, ref unsupported);
        if (!style.Visible) return;
        OfficeTransform transform = ResolveTransform(element, inheritedTransform, viewX, viewY, ref unsupported);
        if (name == "foreignobject") {
            TryAddForeignObject(
                element,
                drawing,
                style,
                references,
                transform,
                viewX,
                viewY,
                maximumElements,
                ref visited,
                ref unsupported);
            return;
        }
        if (name == "image") {
            if (!TryAddEmbeddedSvgImage(element, drawing, style, transform, viewX, viewY)) unsupported++;
            return;
        }
        if (name == "svg") {
            if (!TryAddNestedSvgViewport(
                    element, drawing, style, paintServers, references, transform, viewX, viewY,
                    maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                    ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported)) {
                unsupported++;
            }
            return;
        }
        if (name is "g" or "a" or "switch") {
            bool hasEffects = TryResolveSvgEffects(
                element,
                drawing.Width,
                drawing.Height,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
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
            bool capturesLink = name == "a";
            OfficeDrawing target = hasEffects || capturesLink ? new OfficeDrawing(drawing.Width, drawing.Height) : drawing;
            if (name == "switch") {
                AddFirstSupportedSwitchChild(element, target, style, paintServers, references, transform, viewX, viewY,
                    maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                    ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            } else {
                AddChildren(element, target, style, paintServers, references, transform, viewX, viewY,
                    maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                    ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            }
            if (hasEffects) {
                TryApplySvgFilter(target, filterEffect, transform, maximumElements, ref visited, ref unsupported, out target);
                drawing.AddEffectDrawing(target, OfficeTransform.Identity, blendMode, softMask);
            } else if (capturesLink) {
                drawing.AddDrawingForClippedRendering(target, 0D, 0D, null);
            }
            if (capturesLink) {
                TryAddSvgLink(element, target, drawing, ref unsupported);
            }
            return;
        }
        if (name is "use" or "text") {
            bool hasEffects = TryResolveSvgEffects(
                element,
                drawing.Width,
                drawing.Height,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
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
            OfficeDrawing target = hasEffects ? new OfficeDrawing(drawing.Width, drawing.Height) : drawing;
            if (name == "use") {
                AddReferencedElement(element, target, style, paintServers, references, transform, viewX, viewY,
                    maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                    ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            } else {
                AddText(
                    element,
                    target,
                    style,
                    paintServers,
                    references,
                    transform,
                    viewX,
                    viewY,
                    maximumElements,
                    maximumViewportDimension,
                    maximumViewportPixels,
                    depth,
                    ref visited,
                    ref pathCommands,
                    ref pathCommandLimitExceeded,
                    ref unsupported);
            }
            if (hasEffects) {
                TryApplySvgFilter(target, filterEffect, transform, maximumElements, ref visited, ref unsupported, out target);
                drawing.AddEffectDrawing(target, OfficeTransform.Identity, blendMode, softMask);
            }
            return;
        }

        OfficeDrawingShape? shape = name switch {
            "rect" => CreateRectangle(element, style, viewX, viewY, drawing.Width, drawing.Height, ref unsupported),
            "circle" => CreateCircle(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "ellipse" => CreateEllipse(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "line" => CreateLine(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "polygon" => CreatePolygon(element, style, viewX, viewY, close: true, ref pathCommands, ref pathCommandLimitExceeded),
            "polyline" => CreatePolygon(element, style, viewX, viewY, close: false, ref pathCommands, ref pathCommandLimitExceeded),
            "path" => CreatePath(element, style, viewX, viewY, ref pathCommands, ref pathCommandLimitExceeded),
            _ => null
        };
        if (shape == null) {
            unsupported++;
            return;
        }

        ApplyDeferredPaint(shape.Shape, style, shape.X, shape.Y, drawing.Width, drawing.Height, viewX, viewY, ref unsupported);

        ApplyTransform(shape, transform);

        try {
            bool hasPattern = TryAddSvgPatternFill(
                style.FillPattern,
                shape,
                drawing,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
                maximumElements,
                maximumViewportDimension,
                maximumViewportPixels,
                depth,
                ref visited,
                ref pathCommands,
                ref pathCommandLimitExceeded,
                ref unsupported,
                out OfficeDrawing? patternLayer);
            bool hasStrokePattern = TryAddSvgPatternStroke(
                style.StrokePattern,
                shape,
                drawing,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
                maximumElements,
                maximumViewportDimension,
                maximumViewportPixels,
                depth,
                ref visited,
                ref pathCommands,
                ref pathCommandLimitExceeded,
                ref unsupported,
                out OfficeDrawing? strokePatternLayer);
            bool hasMarkers = TryAddSvgMarkers(
                shape,
                drawing,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
                maximumElements,
                maximumViewportDimension,
                maximumViewportPixels,
                depth,
                ref visited,
                ref pathCommands,
                ref pathCommandLimitExceeded,
                ref unsupported,
                out OfficeDrawing? markerLayer);
            bool hasEffects = TryResolveSvgEffects(
                element,
                drawing.Width,
                drawing.Height,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
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
            if (hasEffects || hasPattern || hasStrokePattern || hasMarkers) {
                var target = new OfficeDrawing(drawing.Width, drawing.Height);
                if (patternLayer != null) target.AddEffectDrawing(patternLayer, OfficeTransform.Identity);
                target.AddShapeForClippedRendering(shape.Shape, shape.X, shape.Y);
                if (strokePatternLayer != null) target.AddEffectDrawing(strokePatternLayer, OfficeTransform.Identity);
                if (markerLayer != null) target.AddEffectDrawing(markerLayer, OfficeTransform.Identity);
                TryApplySvgFilter(target, filterEffect, transform, maximumElements, ref visited, ref unsupported, out target);
                drawing.AddEffectDrawing(target, OfficeTransform.Identity, blendMode, softMask);
            } else {
                drawing.AddShapeForClippedRendering(shape.Shape, shape.X, shape.Y);
            }
        } catch (ArgumentOutOfRangeException) {
            unsupported++;
        }
    }

    private static void AddFirstSupportedSwitchChild(
        XElement element,
        OfficeDrawing drawing,
        SvgPaintContext inherited,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported) {
        foreach (XElement child in element.Elements()) {
            string childName = child.Name.LocalName.ToLowerInvariant();
            if (childName is "title" or "desc" or "metadata" or "defs" or "style"
                or "lineargradient" or "radialgradient" or "pattern" or "stop") {
                continue;
            }
            if (child.Attribute("requiredExtensions") != null || child.Attribute("requiredFeatures") != null) {
                continue;
            }
            if (!IsSupportedSwitchElement(childName)) continue;
            AddElement(child, drawing, inherited, paintServers, references, inheritedTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            return;
        }
    }

    private static bool IsSupportedSwitchElement(string name) => name is
        "svg" or "g" or "a" or "switch" or "foreignobject" or "use" or "text"
        or "image" or "rect" or "circle" or "ellipse" or "line" or "polygon" or "polyline" or "path";

}
