using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryAddRenderedSvgLocalReferenceApplications(
        string? value,
        string expectedTargetName,
        int applications,
        SvgElementReferenceRegistry references,
        int maximumElements,
        ref int elementCount,
        ref int commandCount,
        OfficeTransform transform,
        double viewX,
        double viewY,
        SvgRasterWorkBudget rasterWork) {
        if (applications <= 0) return true;
        for (int application = 0; application < applications; application++) {
            SvgElementReferenceEntryResult result = references.TryEnterLocalDetailed(
                value,
                expectedTargetName,
                out string referenceId,
                out XElement? target);
            if (result is SvgElementReferenceEntryResult.DepthExceeded or SvgElementReferenceEntryResult.Cycle) return false;
            if (result != SvgElementReferenceEntryResult.Entered) return !HasPotentialSvgUrlFunction(value);
            try {
                // Marker viewport, ref-point, orientation, and markerUnits transforms vary per
                // placement. Charge the complete raster viewport before expanding each instance.
                if (!rasterWork.TryChargeFullViewport()) return false;
                rasterWork.EnterConservativePlacement();
                try {
                    if (!TryResolveRenderedSvgAncestorStrokeStyle(target!, out SvgRasterStrokeStyle strokeStyle)
                        || !TryResolveRenderedSvgAncestorTextStyle(target!, out SvgRasterTextStyle textStyle)
                        || !TryAddRenderedSvgExpansion(
                            target!,
                            references,
                            maximumElements,
                            ref elementCount,
                            ref commandCount,
                            transform,
                            viewX,
                            viewY,
                            rasterWork,
                            inheritedStrokeStyle: strokeStyle,
                            inheritedTextStyle: textStyle)) return false;
                } finally {
                    rasterWork.ExitConservativePlacement();
                }
            } finally {
                references.Exit(referenceId);
            }
        }
        return true;
    }

    private static SvgMarkerPlacementCounts CountSvgMarkerPlacements(XElement element) {
        string name = element.Name.LocalName.ToLowerInvariant();
        if (name == "line") return SvgMarkerPlacementCounts.ForOpenVertices(2);
        if (name is "rect" or "circle" or "ellipse") return SvgMarkerPlacementCounts.ForClosedVertices(4);
        if (name is "polygon" or "polyline") {
            if (!TryParseNumberList(
                    ReadRasterProjectedAttribute(element, "points"),
                    MaximumSvgPathCommands * 2,
                    out IReadOnlyList<double> values,
                    out _)) return default;
            int vertices = values.Count / 2;
            return name == "polygon"
                ? SvgMarkerPlacementCounts.ForClosedVertices(vertices)
                : SvgMarkerPlacementCounts.ForOpenVertices(vertices);
        }
        if (name != "path") return default;
        _ = OfficeSvgPathDataParser.TryParse(
            ReadRasterProjectedAttribute(element, "d"),
            MaximumSvgPathCommands,
            out IReadOnlyList<OfficePathCommand> commands,
            out _);
        int start = 0;
        int mid = 0;
        int end = 0;
        int verticesInSubpath = 0;
        bool closed = false;
        foreach (OfficePathCommand command in commands) {
            if (command.Kind == OfficePathCommandKind.MoveTo) {
                AddSvgSubpathMarkerPlacements(verticesInSubpath, closed, ref start, ref mid, ref end);
                verticesInSubpath = 1;
                closed = false;
            } else if (command.Kind == OfficePathCommandKind.Close) {
                closed = true;
            } else if (verticesInSubpath > 0) {
                verticesInSubpath++;
            }
        }
        AddSvgSubpathMarkerPlacements(verticesInSubpath, closed, ref start, ref mid, ref end);
        return new SvgMarkerPlacementCounts(start, mid, end);
    }

    private static void AddSvgSubpathMarkerPlacements(
        int vertices,
        bool closed,
        ref int start,
        ref int mid,
        ref int end) {
        if (vertices <= 0) return;
        start++;
        end++;
        mid += Math.Max(0, vertices - 2);
        if (closed && vertices > 1) mid++;
    }

    private readonly struct SvgMarkerPlacementCounts {
        internal SvgMarkerPlacementCounts(int start, int mid, int end) {
            Start = start;
            Mid = mid;
            End = end;
        }

        internal int Start { get; }
        internal int Mid { get; }
        internal int End { get; }
        internal bool HasAny => Start > 0 || Mid > 0 || End > 0;

        internal static SvgMarkerPlacementCounts ForOpenVertices(int vertices) => vertices <= 0
            ? default
            : new SvgMarkerPlacementCounts(1, Math.Max(0, vertices - 2), 1);

        internal static SvgMarkerPlacementCounts ForClosedVertices(int vertices) => vertices <= 0
            ? default
            : new SvgMarkerPlacementCounts(1, Math.Max(0, vertices - 1), 1);
    }
}
