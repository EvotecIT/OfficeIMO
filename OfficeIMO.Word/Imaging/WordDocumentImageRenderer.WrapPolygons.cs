using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double WordWrapPolygonCoordinateExtent = 21600D;
        private const double WordWrapPolygonFullEdgeThreshold = 21000D;

        private static bool TryCreateAuthoredWrapPolygonTextExclusion(
            DW.Anchor? anchor,
            double left,
            double top,
            double right,
            double bottom,
            out IReadOnlyList<OfficePoint> polygon) {
            polygon = Array.Empty<OfficePoint>();
            if (anchor == null || right <= left || bottom <= top) {
                return false;
            }

            DW.WrapPolygon? wrapPolygon =
                anchor.GetFirstChild<DW.WrapTight>()?.GetFirstChild<DW.WrapPolygon>() ??
                anchor.GetFirstChild<DW.WrapThrough>()?.GetFirstChild<DW.WrapPolygon>();
            if (wrapPolygon == null) {
                return false;
            }

            var rawPoints = new List<(long X, long Y)>();
            DW.StartPoint? start = wrapPolygon.GetFirstChild<DW.StartPoint>();
            if (start?.X?.Value is long startX && start.Y?.Value is long startY) {
                rawPoints.Add((startX, startY));
            }

            foreach (DW.LineTo line in wrapPolygon.Elements<DW.LineTo>()) {
                if (line.X?.Value is long x && line.Y?.Value is long y) {
                    rawPoints.Add((x, y));
                }
            }

            if (rawPoints.Count < 3 || CoversFullRectangle(rawPoints)) {
                return false;
            }

            double width = right - left;
            double height = bottom - top;
            polygon = rawPoints
                .Select(point => new OfficePoint(
                    left + (ClampWrapPolygonCoordinate(point.X) / WordWrapPolygonCoordinateExtent * width),
                    top + (ClampWrapPolygonCoordinate(point.Y) / WordWrapPolygonCoordinateExtent * height)))
                .ToArray();
            return true;
        }

        private static bool TryCreateTransparentImageWrapPolygon(
            byte[] bytes,
            OfficeImageProjection projection,
            out IReadOnlyList<OfficePoint> polygon) {
            polygon = Array.Empty<OfficePoint>();
            if (projection.HasTransform ||
                projection.Width <= 0D ||
                projection.Height <= 0D ||
                !OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? raster) ||
                raster == null) {
                return false;
            }

            int sourceLeft = ClampPixelIndex((int)Math.Floor(projection.SourceLeft * raster.Width), raster.Width);
            int sourceTop = ClampPixelIndex((int)Math.Floor(projection.SourceTop * raster.Height), raster.Height);
            int sourceRight = ClampPixelEdge((int)Math.Ceiling((projection.SourceLeft + projection.SourceWidth) * raster.Width), raster.Width);
            int sourceBottom = ClampPixelEdge((int)Math.Ceiling((projection.SourceTop + projection.SourceHeight) * raster.Height), raster.Height);
            if (sourceRight <= sourceLeft || sourceBottom <= sourceTop) {
                return false;
            }

            var leftEdge = new List<OfficePoint>();
            var rightEdge = new List<OfficePoint>();
            bool foundOpaqueRow = false;
            bool foundGapAfterOpaque = false;
            bool sawTransparentPixel = false;
            int previousLeft = 0;
            int previousRight = 0;
            double previousBottom = 0D;

            for (int y = sourceTop; y < sourceBottom; y++) {
                int rowLeft = int.MaxValue;
                int rowRight = int.MinValue;
                for (int x = sourceLeft; x < sourceRight; x++) {
                    if (raster.GetPixel(x, y).A > 0) {
                        rowLeft = Math.Min(rowLeft, x);
                        rowRight = Math.Max(rowRight, x + 1);
                    } else {
                        sawTransparentPixel = true;
                    }
                }

                if (rowRight <= rowLeft) {
                    if (foundOpaqueRow) {
                        foundGapAfterOpaque = true;
                    }

                    continue;
                }

                if (foundGapAfterOpaque) {
                    return false;
                }

                if (rowLeft > sourceLeft || rowRight < sourceRight) {
                    sawTransparentPixel = true;
                }

                double rowTop = MapSourceYToDestination(y, sourceTop, sourceBottom, projection);
                double rowBottom = MapSourceYToDestination(y + 1, sourceTop, sourceBottom, projection);
                if (!foundOpaqueRow) {
                    leftEdge.Add(new OfficePoint(MapSourceXToDestination(rowLeft, sourceLeft, sourceRight, projection), rowTop));
                    rightEdge.Add(new OfficePoint(MapSourceXToDestination(rowRight, sourceLeft, sourceRight, projection), rowTop));
                } else {
                    if (rowLeft != previousLeft) {
                        leftEdge.Add(new OfficePoint(MapSourceXToDestination(previousLeft, sourceLeft, sourceRight, projection), rowTop));
                        leftEdge.Add(new OfficePoint(MapSourceXToDestination(rowLeft, sourceLeft, sourceRight, projection), rowTop));
                    }

                    if (rowRight != previousRight) {
                        rightEdge.Add(new OfficePoint(MapSourceXToDestination(previousRight, sourceLeft, sourceRight, projection), rowTop));
                        rightEdge.Add(new OfficePoint(MapSourceXToDestination(rowRight, sourceLeft, sourceRight, projection), rowTop));
                    }
                }

                foundOpaqueRow = true;
                previousLeft = rowLeft;
                previousRight = rowRight;
                previousBottom = rowBottom;
            }

            if (!foundOpaqueRow || !sawTransparentPixel) {
                return false;
            }

            leftEdge.Add(new OfficePoint(MapSourceXToDestination(previousLeft, sourceLeft, sourceRight, projection), previousBottom));
            rightEdge.Add(new OfficePoint(MapSourceXToDestination(previousRight, sourceLeft, sourceRight, projection), previousBottom));

            var points = new List<OfficePoint>(leftEdge.Count + rightEdge.Count);
            points.AddRange(leftEdge);
            for (int index = rightEdge.Count - 1; index >= 0; index--) {
                points.Add(rightEdge[index]);
            }

            polygon = points;
            return polygon.Count >= 3;
        }

        private static double ClampWrapPolygonCoordinate(long value) =>
            Math.Min(WordWrapPolygonCoordinateExtent, Math.Max(0D, value));

        private static bool CoversFullRectangle(IReadOnlyList<(long X, long Y)> points) {
            long minX = points.Min(point => point.X);
            long minY = points.Min(point => point.Y);
            long maxX = points.Max(point => point.X);
            long maxY = points.Max(point => point.Y);
            return minX <= 0L &&
                   minY <= 0L &&
                   maxX >= WordWrapPolygonFullEdgeThreshold &&
                   maxY >= WordWrapPolygonFullEdgeThreshold;
        }

        private static int ClampPixelIndex(int value, int size) =>
            Math.Min(Math.Max(0, value), Math.Max(0, size - 1));

        private static int ClampPixelEdge(int value, int size) =>
            Math.Min(Math.Max(1, value), size);

        private static double MapSourceXToDestination(int x, int sourceLeft, int sourceRight, OfficeImageProjection projection) =>
            projection.X + (((double)(x - sourceLeft) / Math.Max(1, sourceRight - sourceLeft)) * projection.Width);

        private static double MapSourceYToDestination(int y, int sourceTop, int sourceBottom, OfficeImageProjection projection) =>
            projection.Y + (((double)(y - sourceTop) / Math.Max(1, sourceBottom - sourceTop)) * projection.Height);
    }
}
