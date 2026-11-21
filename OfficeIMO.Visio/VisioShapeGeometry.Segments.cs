namespace OfficeIMO.Visio {
    /// <summary>One exact outline segment in local shape coordinates: a line or a cubic Bezier ending at <see cref="End"/>.</summary>
    internal readonly struct VisioGeometrySegment {
        private VisioGeometrySegment(bool isCubic, (double X, double Y) control1, (double X, double Y) control2, (double X, double Y) end) {
            IsCubic = isCubic;
            Control1 = control1;
            Control2 = control2;
            End = end;
        }

        internal bool IsCubic { get; }

        internal (double X, double Y) Control1 { get; }

        internal (double X, double Y) Control2 { get; }

        internal (double X, double Y) End { get; }

        internal static VisioGeometrySegment Line((double X, double Y) end) => new(false, default, default, end);

        internal static VisioGeometrySegment Cubic((double X, double Y) control1, (double X, double Y) control2, (double X, double Y) end) =>
            new(true, control1, control2, end);

        internal VisioGeometrySegment Map(Func<(double X, double Y), (double X, double Y)> map) =>
            new(IsCubic, IsCubic ? map(Control1) : default, IsCubic ? map(Control2) : default, map(End));
    }

    internal static partial class VisioShapeGeometry {
        /// <summary>
        /// Keeps vector consumers' exact segments aligned with the flattened points that raster, hit-testing and
        /// layout consumers use. Rows without an exact form are recorded as the lines between their flattened points.
        /// </summary>
        private sealed class SegmentRecorder {
            private int _synced;

            internal List<VisioGeometrySegment> Segments { get; private set; } = new();

            internal bool HasCurves { get; private set; }

            internal bool Flattened { get; private set; }

            internal void Sync(List<(double X, double Y)> points) {
                for (; _synced < points.Count; _synced++) Segments.Add(VisioGeometrySegment.Line(points[_synced]));
            }

            internal void AddExact(List<VisioGeometrySegment> cubics, List<(double X, double Y)> points) {
                if (cubics.Count == 0) return;
                Segments.AddRange(cubics);
                HasCurves = true;
                _synced = points.Count;
            }

            internal void MarkFlattened() => Flattened = true;

            internal void Reset() {
                Segments = new List<VisioGeometrySegment>();
                _synced = 0;
                HasCurves = false;
                Flattened = false;
            }
        }

        /// <summary>Approximates an affine image of a circular arc with at most quarter-turn cubic Bezier segments.</summary>
        private static void AppendArcCubics(
            List<VisioGeometrySegment> cubics,
            Func<double, (double X, double Y)> point,
            Func<double, (double X, double Y)> derivative,
            double startAngle,
            double sweep,
            (double X, double Y)? exactEnd) {
            if (!IsFinite(sweep) || Math.Abs(sweep) <= 1e-12) return;
            int count = Math.Max(1, (int)Math.Ceiling((Math.Abs(sweep) / (Math.PI / 2D)) - 1e-9));
            double step = sweep / count;
            double scale = 4D / 3D * Math.Tan(step / 4D);
            for (int index = 0; index < count; index++) {
                double from = startAngle + (step * index);
                double to = from + step;
                (double X, double Y) start = point(from);
                (double X, double Y) end = index == count - 1 && exactEnd.HasValue ? exactEnd.Value : point(to);
                (double X, double Y) startTangent = derivative(from);
                (double X, double Y) endTangent = derivative(to);
                cubics.Add(VisioGeometrySegment.Cubic(
                    (start.X + (scale * startTangent.X), start.Y + (scale * startTangent.Y)),
                    (end.X - (scale * endTangent.X), end.Y - (scale * endTangent.Y)),
                    end));
            }
        }

        /// <summary>Builds the named ellipse fallback with exact cubic segments alongside its flattened points.</summary>
        internal static VisioShapeGeometryPath CreateBuiltinEllipsePath(VisioShape shape) {
            var points = new List<(double X, double Y)>();
            var cubics = new List<VisioGeometrySegment>();
            double rx = shape.WidthInches / 2D, ry = shape.HeightInches / 2D;
            AppendEllipse(points, (rx, ry), (rx * 2D, ry), (rx, ry * 2D), cubics);
            if (points.Count < 3) return new VisioShapeGeometryPath(GetBuiltinClosedPath(shape, "rectangle"), false, false, true, 0);
            cubics.Insert(0, VisioGeometrySegment.Line(points[0]));
            return new VisioShapeGeometryPath(points, false, false, true, 0, cubics);
        }

        private static List<VisioGeometrySegment> CubicFromQuadratic((double X, double Y) start, (double X, double Y) control, (double X, double Y) end) =>
            new() {
                VisioGeometrySegment.Cubic(
                    (start.X + (2D / 3D * (control.X - start.X)), start.Y + (2D / 3D * (control.Y - start.Y))),
                    (end.X + (2D / 3D * (control.X - end.X)), end.Y + (2D / 3D * (control.Y - end.Y))),
                    end)
            };
    }
}
