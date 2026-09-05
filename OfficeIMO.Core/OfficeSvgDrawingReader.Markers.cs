using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumRenderedSvgMarkers = 4096;

    private static bool TryAddSvgMarkers(
        OfficeDrawingShape shape,
        OfficeDrawing drawing,
        SvgPaintContext style,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform elementTransform,
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
        out OfficeDrawing? markerLayer) {
        markerLayer = null;
        if (style.MarkerStart == null && style.MarkerMid == null && style.MarkerEnd == null) return false;
        IReadOnlyList<SvgMarkerPlacement> placements = ResolveSvgMarkerPlacements(shape);
        if (placements.Count == 0) {
            unsupported++;
            return false;
        }
        if (placements.Count > MaximumRenderedSvgMarkers) {
            unsupported++;
            return false;
        }

        var layer = new OfficeDrawing(drawing.Width, drawing.Height);
        bool rendered = false;
        foreach (SvgMarkerPlacement placement in placements) {
            string? reference = placement.Kind switch {
                SvgMarkerPlacementKind.Start => style.MarkerStart,
                SvgMarkerPlacementKind.Mid => style.MarkerMid,
                _ => style.MarkerEnd
            };
            if (reference == null) continue;
            if (!TryRenderSvgMarker(reference, placement, elementTransform, layer, style, paintServers, references,
                    maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                    ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported)) continue;
            rendered = true;
        }
        if (!rendered) return false;
        markerLayer = layer;
        return true;
    }

    private static bool TryRenderSvgMarker(
        string reference,
        SvgMarkerPlacement placement,
        OfficeTransform elementTransform,
        OfficeDrawing layer,
        SvgPaintContext inheritedStyle,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported) {
        SvgElementReferenceEntryResult entry = references.TryEnterLocalDetailed(reference, "marker", out string id, out XElement? marker);
        if (entry != SvgElementReferenceEntryResult.Entered) {
            unsupported++;
            return false;
        }
        try {
            if (!TrySvgLength(marker!.Attribute("markerWidth")?.Value ?? "3", out double markerWidth)
                || !TrySvgLength(marker.Attribute("markerHeight")?.Value ?? "3", out double markerHeight)
                || markerWidth <= 0D
                || markerHeight <= 0D) {
                unsupported++;
                return false;
            }
            string markerUnits = marker.Attribute("markerUnits")?.Value.Trim() ?? "strokeWidth";
            if (markerUnits.Equals("strokeWidth", StringComparison.OrdinalIgnoreCase)) {
                markerWidth *= Math.Max(0.01D, inheritedStyle.StrokeWidth);
                markerHeight *= Math.Max(0.01D, inheritedStyle.StrokeWidth);
            } else if (!markerUnits.Equals("userSpaceOnUse", StringComparison.OrdinalIgnoreCase)) {
                unsupported++;
                return false;
            }
            if (!IsSupportedSvgViewport(markerWidth, markerHeight, maximumViewportDimension, maximumViewportPixels)) {
                unsupported++;
                return false;
            }

            IReadOnlyList<double> viewBox;
            if (marker.Attribute("viewBox") == null) {
                viewBox = new[] { 0D, 0D, markerWidth, markerHeight };
            } else if (!TryParseNumberList(marker.Attribute("viewBox")?.Value, out viewBox)
                       || viewBox.Count != 4
                       || viewBox[2] <= 0D
                       || viewBox[3] <= 0D) {
                unsupported++;
                return false;
            }
            if (!TryParsePreserveAspectRatio(marker.Attribute("preserveAspectRatio")?.Value, out SvgAspectAlignment alignment, out bool slice)
                || !TryMarkerCoordinate(marker.Attribute("refX")?.Value, viewBox[2], 0D, out double refX)
                || !TryMarkerCoordinate(marker.Attribute("refY")?.Value, viewBox[3], 0D, out double refY)
                || !TryResolveMarkerAngle(marker.Attribute("orient")?.Value, placement, out double angle)) {
                unsupported++;
                return false;
            }

            var scene = new OfficeDrawing(viewBox[2], viewBox[3]);
            SvgPaintContext markerStyle = ResolvePaintContext(marker, inheritedStyle, paintServers, ref unsupported);
            markerStyle.MarkerStart = null;
            markerStyle.MarkerMid = null;
            markerStyle.MarkerEnd = null;
            OfficeTransform markerContentTransform = ResolveTransform(marker, OfficeTransform.Identity, viewBox[0], viewBox[1], ref unsupported);
            AddChildren(marker, scene, markerStyle, paintServers, references, markerContentTransform, viewBox[0], viewBox[1],
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            if (scene.Elements.Count == 0) return false;

            OfficeTransform viewportTransform = ResolveViewportTransform(viewBox[2], viewBox[3], markerWidth, markerHeight, alignment, slice);
            OfficePoint refPoint = viewportTransform.TransformPoint(new OfficePoint(refX - viewBox[0], refY - viewBox[1]));
            OfficeTransform placementTransform = viewportTransform
                .Then(OfficeTransform.Translate(-refPoint.X, -refPoint.Y))
                .Then(OfficeTransform.RotateDegrees(angle))
                .Then(OfficeTransform.Translate(placement.Point.X, placement.Point.Y))
                .Then(elementTransform);
            layer.AddEffectDrawing(scene, placementTransform);
            return true;
        } finally {
            references.Exit(id);
        }
    }

    private static bool TryMarkerCoordinate(string? text, double extent, double fallback, out double value) {
        value = fallback;
        return string.IsNullOrWhiteSpace(text) || TryViewportLength(text, extent, out value, out _);
    }

    private static bool TryResolveMarkerAngle(string? orient, SvgMarkerPlacement placement, out double angle) {
        angle = placement.AngleDegrees;
        if (string.IsNullOrWhiteSpace(orient) || orient!.Trim().Equals("0", StringComparison.OrdinalIgnoreCase)) {
            angle = 0D;
            return true;
        }
        string normalized = orient.Trim();
        if (normalized.Equals("auto", StringComparison.OrdinalIgnoreCase)) return true;
        if (normalized.Equals("auto-start-reverse", StringComparison.OrdinalIgnoreCase)) {
            if (placement.Kind == SvgMarkerPlacementKind.Start) angle += 180D;
            return true;
        }
        if (normalized.EndsWith("deg", StringComparison.OrdinalIgnoreCase)) normalized = normalized.Substring(0, normalized.Length - 3).Trim();
        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out angle)
            && !double.IsNaN(angle)
            && !double.IsInfinity(angle);
    }

    private static IReadOnlyList<SvgMarkerPlacement> ResolveSvgMarkerPlacements(OfficeDrawingShape shape) {
        var contours = new List<OfficeFlattenedPathContour>();
        if (shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.Points.Count >= 2) {
            contours.Add(new OfficeFlattenedPathContour(shape.Shape.Points, closed: false));
        } else if (shape.Shape.Kind == OfficeShapeKind.Polygon && shape.Shape.Points.Count >= 2) {
            contours.Add(new OfficeFlattenedPathContour(shape.Shape.Points, closed: true));
        } else if (shape.Shape.Kind == OfficeShapeKind.Path) {
            contours.AddRange(OfficePathFlattener.Flatten(shape.Shape.PathCommands, 0D, 0D, 1D));
        }

        var result = new List<SvgMarkerPlacement>();
        foreach (OfficeFlattenedPathContour contour in contours) {
            if (contour.Points.Count < 2) continue;
            var points = contour.Points
                .Select(point => new OfficePoint(shape.X + point.X, shape.Y + point.Y))
                .ToArray();
            result.Add(new SvgMarkerPlacement(points[0], ResolveSegmentAngle(points[0], points[1]), SvgMarkerPlacementKind.Start));
            if (contour.Closed) {
                for (int index = 1; index < points.Length; index++) {
                    OfficePoint next = index == points.Length - 1 ? points[0] : points[index + 1];
                    result.Add(new SvgMarkerPlacement(points[index], ResolveBisectedAngle(points[index - 1], points[index], next), SvgMarkerPlacementKind.Mid));
                }
                result.Add(new SvgMarkerPlacement(points[0], ResolveSegmentAngle(points[points.Length - 1], points[0]), SvgMarkerPlacementKind.End));
            } else {
                for (int index = 1; index < points.Length - 1; index++) {
                    result.Add(new SvgMarkerPlacement(points[index], ResolveBisectedAngle(points[index - 1], points[index], points[index + 1]), SvgMarkerPlacementKind.Mid));
                }
                result.Add(new SvgMarkerPlacement(points[points.Length - 1], ResolveSegmentAngle(points[points.Length - 2], points[points.Length - 1]), SvgMarkerPlacementKind.End));
            }
        }
        return result;
    }

    private static double ResolveSegmentAngle(OfficePoint start, OfficePoint end) =>
        Math.Atan2(end.Y - start.Y, end.X - start.X) * 180D / Math.PI;

    private static double ResolveBisectedAngle(OfficePoint previous, OfficePoint point, OfficePoint next) {
        double incoming = ResolveSegmentAngle(previous, point);
        double outgoing = ResolveSegmentAngle(point, next);
        double difference = Math.IEEERemainder(outgoing - incoming, 360D);
        return incoming + (difference / 2D);
    }

    private enum SvgMarkerPlacementKind {
        Start,
        Mid,
        End
    }

    private readonly struct SvgMarkerPlacement {
        internal SvgMarkerPlacement(OfficePoint point, double angleDegrees, SvgMarkerPlacementKind kind) {
            Point = point;
            AngleDegrees = angleDegrees;
            Kind = kind;
        }

        internal OfficePoint Point { get; }
        internal double AngleDegrees { get; }
        internal SvgMarkerPlacementKind Kind { get; }
    }
}
