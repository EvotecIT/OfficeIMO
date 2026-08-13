using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly partial struct PdfPageClipPath {
    private static bool IsSegmentWithinFilledArea(
        List<List<OfficePoint>> contours,
        OfficeFillRule fillRule,
        OfficePoint start,
        OfficePoint end) {
        var parameters = new List<double> { 0D, 1D };
        for (int contourIndex = 0; contourIndex < contours.Count; contourIndex++) {
            List<OfficePoint> contour = contours[contourIndex];
            for (int pointIndex = 0; pointIndex < contour.Count; pointIndex++) {
                AddIntersectionParameters(
                    start,
                    end,
                    contour[pointIndex],
                    contour[(pointIndex + 1) % contour.Count],
                    parameters);
            }
        }
        parameters.Sort();
        for (int index = 1; index < parameters.Count; index++) {
            double lower = parameters[index - 1];
            double upper = parameters[index];
            if (upper - lower <= 0.000001D) continue;
            double t = (lower + upper) / 2D;
            var sample = new OfficePoint(
                start.X + ((end.X - start.X) * t),
                start.Y + ((end.Y - start.Y) * t));
            if (!ContainsFilledPoint(contours, fillRule, sample)) return false;
        }
        return true;
    }

    private static void AddIntersectionParameters(
        OfficePoint firstStart,
        OfficePoint firstEnd,
        OfficePoint secondStart,
        OfficePoint secondEnd,
        List<double> parameters) {
        double firstX = firstEnd.X - firstStart.X;
        double firstY = firstEnd.Y - firstStart.Y;
        double secondX = secondEnd.X - secondStart.X;
        double secondY = secondEnd.Y - secondStart.Y;
        double denominator = (firstX * secondY) - (firstY * secondX);
        if (Math.Abs(denominator) > 0.000001D) {
            double offsetX = secondStart.X - firstStart.X;
            double offsetY = secondStart.Y - firstStart.Y;
            double t = ((offsetX * secondY) - (offsetY * secondX)) / denominator;
            double u = ((offsetX * firstY) - (offsetY * firstX)) / denominator;
            if (t >= -0.000001D && t <= 1.000001D && u >= -0.000001D && u <= 1.000001D) {
                AddDistinctParameter(parameters, Math.Max(0D, Math.Min(1D, t)));
            }
            return;
        }
        if (Math.Abs(Cross(firstStart, firstEnd, secondStart)) > 0.001D) return;
        AddProjectedParameter(firstStart, firstEnd, secondStart, parameters);
        AddProjectedParameter(firstStart, firstEnd, secondEnd, parameters);
    }

    private static void AddProjectedParameter(
        OfficePoint start,
        OfficePoint end,
        OfficePoint point,
        List<double> parameters) {
        double x = end.X - start.X;
        double y = end.Y - start.Y;
        double lengthSquared = (x * x) + (y * y);
        if (lengthSquared <= 0.000000000001D || !IsPointOnSegment(point, start, end)) return;
        double t = (((point.X - start.X) * x) + ((point.Y - start.Y) * y)) / lengthSquared;
        AddDistinctParameter(parameters, Math.Max(0D, Math.Min(1D, t)));
    }

    private static void AddDistinctParameter(List<double> parameters, double value) {
        for (int index = 0; index < parameters.Count; index++) {
            if (Math.Abs(parameters[index] - value) <= 0.000001D) return;
        }
        parameters.Add(value);
    }
}
