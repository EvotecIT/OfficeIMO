using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private readonly record struct SceneVector(double X, double Y, double Z) {
        internal double Length => Math.Sqrt(Dot(this, this));
        internal double Largest => Math.Max(X, Math.Max(Y, Z));
        internal double Smallest => Math.Min(X, Math.Min(Y, Z));

        public static SceneVector operator +(SceneVector a, SceneVector b) => new(a.X + b.X, a.Y + b.Y, a.Z + b.Z);
        public static SceneVector operator -(SceneVector a, SceneVector b) => new(a.X - b.X, a.Y - b.Y, a.Z - b.Z);
        public static SceneVector operator *(SceneVector value, double factor) => new(value.X * factor, value.Y * factor, value.Z * factor);

        internal static SceneVector Min(SceneVector a, SceneVector b) => new(Math.Min(a.X, b.X), Math.Min(a.Y, b.Y), Math.Min(a.Z, b.Z));
        internal static SceneVector Max(SceneVector a, SceneVector b) => new(Math.Max(a.X, b.X), Math.Max(a.Y, b.Y), Math.Max(a.Z, b.Z));
        internal static double Dot(SceneVector a, SceneVector b) => a.X * b.X + a.Y * b.Y + a.Z * b.Z;
        internal static SceneVector Cross(SceneVector a, SceneVector b) => new(a.Y * b.Z - a.Z * b.Y,
            a.Z * b.X - a.X * b.Z, a.X * b.Y - a.Y * b.X);
    }

    private readonly record struct SceneCamera(SceneVector Target, SceneVector Forward, SceneVector Right,
        SceneVector Up, double Distance, bool Orthographic) {
        internal static bool TryCreate(SceneVector eye, SceneVector center, SceneVector up, bool orthographic, out SceneCamera result) {
            result = default;
            SceneVector sight = center - eye;
            double distance = sight.Length;
            double upLength = up.Length;
            if (distance < 0.3 || upLength < 1e-6) return false;
            SceneVector forward = sight * (1 / distance);
            SceneVector right = SceneVector.Cross(forward, up * (1 / upLength));
            double rightLength = right.Length;
            if (rightLength < 1e-6) return false;
            right *= 1 / rightLength;
            result = new SceneCamera(center, forward, right, SceneVector.Cross(right, forward), distance, orthographic);
            return true;
        }

        internal bool TryProject(SceneVector point, out SceneProjection projection) {
            SceneVector relative = point - Target;
            double depth = Distance + SceneVector.Dot(relative, Forward);
            if (!Orthographic && depth < 0.12) {
                projection = default;
                return false;
            }
            double scale = Orthographic ? 1 : Distance / depth;
            projection = new SceneProjection(SceneVector.Dot(relative, Right) * scale,
                -SceneVector.Dot(relative, Up) * scale, depth);
            return true;
        }
    }

    private readonly record struct SceneProjection(double X, double Y, double Depth);

    private readonly record struct SceneDomain(SceneVector Min, SceneVector Max, SceneVector Extent) {
        internal SceneVector Normalize(SceneVector point) {
            SceneVector spans = Max - Min;
            return new SceneVector(((point.X - Min.X) / spans.X - 0.5) * Extent.X,
                ((point.Y - Min.Y) / spans.Y - 0.5) * Extent.Y,
                ((point.Z - Min.Z) / spans.Z - 0.5) * Extent.Z);
        }

        internal SceneVector Corner(int index) => new((index & 1) == 0 ? Min.X : Max.X,
            (index & 2) == 0 ? Min.Y : Max.Y, (index & 4) == 0 ? Min.Z : Max.Z);

        internal double AxisValue(int axis, double t) => axis switch {
            0 => Min.X + (Max.X - Min.X) * t,
            1 => Min.Y + (Max.Y - Min.Y) * t,
            _ => Min.Z + (Max.Z - Min.Z) * t
        };

        internal static SceneDomain Create(SceneLayout layout, IReadOnlyList<SceneTrace> traces) {
            var min = traces[0].Points[0];
            var max = min;
            foreach (SceneTrace trace in traces) {
                foreach (SceneVector point in trace.Points) {
                    min = SceneVector.Min(min, point);
                    max = SceneVector.Max(max, point);
                }
            }

            double padding = (max - min).Largest * 0.05;
            (double xMin, double xMax) = ExpandSceneAxis(min.X, max.X, padding);
            (double yMin, double yMax) = ExpandSceneAxis(min.Y, max.Y, padding);
            (double zMin, double zMax) = ExpandSceneAxis(min.Z, max.Z, padding);
            min = new SceneVector(xMin, yMin, zMin);
            max = new SceneVector(xMax, yMax, zMax);

            SceneVector spans = max - min;
            double largest = spans.Largest;
            SceneVector extent = layout.Aspect switch {
                SceneAspect.Manual => layout.Ratio * (1 / layout.Ratio.Largest),
                SceneAspect.Data => new SceneVector(spans.X / largest, spans.Y / largest, spans.Z / largest),
                SceneAspect.Auto when largest <= 4 * spans.Smallest => new SceneVector(spans.X / largest, spans.Y / largest, spans.Z / largest),
                _ => new SceneVector(1, 1, 1)
            };
            return new SceneDomain(min, max, extent);
        }
    }

    private readonly record struct SceneView(SceneCamera Camera, SceneDomain Domain, double Scale, double CenterX, double CenterY) {
        internal bool TryProject(SceneVector point, out SceneProjection projection) =>
            Camera.TryProject(Domain.Normalize(point), out projection);

        internal OfficePoint Place(SceneProjection point) => new(CenterX + point.X * Scale, CenterY + point.Y * Scale);

        internal static bool TryCreate(SceneLayout layout, SceneDomain domain, double rightReserve, double bottomReserve,
            out SceneView result, out SceneProjection[] corners) {
            result = default;
            corners = new SceneProjection[8];
            double halfWidth = 0.04;
            double halfHeight = 0.04;
            for (int index = 0; index < corners.Length; index++) {
                if (!layout.Camera.TryProject(domain.Normalize(domain.Corner(index)), out SceneProjection point)) return false;
                corners[index] = point;
                // Fit around the camera target, retaining the requested center offset.
                halfWidth = Math.Max(halfWidth, Math.Abs(point.X));
                halfHeight = Math.Max(halfHeight, Math.Abs(point.Y));
            }

            double left = 48;
            double right = layout.Width - rightReserve - 48;
            double top = string.IsNullOrWhiteSpace(layout.Title) ? 31 : 49;
            double bottom = layout.Height - bottomReserve - 39;
            double scale = Math.Min((right - left) / (2 * halfWidth), (bottom - top) / (2 * halfHeight));
            if (scale <= 0) return false;
            result = new SceneView(layout.Camera, domain, scale, (left + right) / 2, (top + bottom) / 2);
            return true;
        }
    }

    private static (double Min, double Max) ExpandSceneAxis(double min, double max, double padding) {
        // Preserve precision for small ranges near zero and for close observations at larger offsets.
        if (min != max) return (min, max);
        double radius = Math.Max(Math.Max(padding, Math.Abs(min) * 1e-9), 1e-12);
        if (padding == 0) radius = Math.Max(radius, 0.5);
        return (min - radius, max + radius);
    }

    private static OfficeColor InterpolateSceneColor(OfficeColor first, OfficeColor second, double ratio) =>
        OfficeColor.FromRgb(InterpolateSceneChannel(first.R, second.R, ratio),
            InterpolateSceneChannel(first.G, second.G, ratio), InterpolateSceneChannel(first.B, second.B, ratio));

    private static byte InterpolateSceneChannel(byte a, byte b, double t) => (byte)Math.Round(a + (b - a) * t);
}
