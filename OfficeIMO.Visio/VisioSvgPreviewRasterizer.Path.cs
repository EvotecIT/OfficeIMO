using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool TryParsePath(string? data, out List<SvgPathContour> contours, double pixelsPerUnit = 1D) {
            contours = new List<SvgPathContour>();
            if (!OfficeSvgPathDataParser.TryParse(data, 100000, out IReadOnlyList<OfficePathCommand> commands, out _)) return false;
            try {
                foreach (OfficeFlattenedPathContour contour in OfficePathFlattener.Flatten(commands, 0, 0, 1, pixelsPerUnit: pixelsPerUnit)) {
                    var points = new List<(double X, double Y)>(contour.Points.Count);
                    foreach (OfficePoint point in contour.Points) points.Add((point.X, point.Y));
                    contours.Add(new SvgPathContour(points, contour.Closed));
                }
            } catch (InvalidOperationException) {
                contours.Clear();
                return false;
            }
            return contours.Count > 0;
        }

        private readonly struct SvgPathContour {
            internal SvgPathContour(List<(double X, double Y)> points, bool isClosed) { Points = points; IsClosed = isClosed; }
            internal List<(double X, double Y)> Points { get; }
            internal bool IsClosed { get; }
        }
    }
}
