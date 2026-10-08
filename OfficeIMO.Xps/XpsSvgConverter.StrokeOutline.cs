using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private readonly Dictionary<XElement, string> _strokeOutlines = new();
    private double _strokeResolution = 8D;

    private string StrokeOutline(string path, double thickness, double miterLimit, double[] dashes,
        double offset, string startCap, string endCap, string dashCap, string join, BrushRegion region, IReadOnlyList<StrokeFigure>? figures) {
        if (!OfficeSvgPathDataParser.TryParse(path, 100000, out var commands, out _, allowEmptyGeometry: true))
            throw new InvalidDataException("Invalid XPS stroke geometry.");
        var options = new OfficeStrokeOutlineOptions {
            StartCap = NativeCap(startCap), EndCap = NativeCap(endCap), DashCap = NativeCap(dashCap),
            ClipMiter = true, DegenerateMiterLimitOne = true, CancellationToken = _token,
            ChargePoints = count => {
                if (count > 1_000_000 - _points) throw new InvalidDataException("XPS stroke outline point budget exceeded.");
                _points += count;
            }
        };
        try {
            var contours = new List<OfficeFlattenedPathContour>();
            if (figures == null) contours.AddRange(OfficePathFlattener.FlattenNativeStroke(commands, _strokeResolution));
            else foreach (var figure in figures) {
                if (!OfficeSvgPathDataParser.TryParse(figure.Path, 100000, out var figureCommands, out _, allowEmptyGeometry: true))
                    throw new InvalidDataException("Invalid XPS stroke figure.");
                foreach (var contour in OfficePathFlattener.FlattenNativeStroke(figureCommands, _strokeResolution))
                    contours.Add(new OfficeFlattenedPathContour(contour.Points, contour.Closed, figure.UseStartCap, figure.UseEndCap));
            }
            var outlines = OfficeStrokeGeometry.CreateNative(contours, thickness,
                join == "Round" ? OfficeStrokeLineJoin.Round : join == "Bevel" ? OfficeStrokeLineJoin.Bevel : OfficeStrokeLineJoin.Miter,
                miterLimit, dashes, offset, _strokeResolution, region.X, region.Y, region.X + region.Width, region.Y + region.Height, options);
            var data = new StringBuilder();
            foreach (var contour in outlines) {
                _token.ThrowIfCancellationRequested();
                for (int i = 0; i < contour.Count; i++) data.Append(i == 0 ? "M" : "L").Append(N(contour[i].X)).Append(',').Append(N(contour[i].Y));
                data.Append('Z');
                EnsureOutputCapacity(data.Length);
            }
            return data.ToString();
        } catch (InvalidOperationException error) {
            throw new InvalidDataException("XPS stroke outline complexity limit exceeded.", error);
        }
    }

    private static double StrokeResolution(double resolution, string? transform) {
        if (transform != null) {
            if (!OfficeSvgTransformParser.TryParse(transform, out var matrix)) throw new InvalidDataException("Invalid XPS stroke transform.");
            double largest = Math.Max(Math.Max(Math.Abs(matrix.M11), Math.Abs(matrix.M12)), Math.Max(Math.Abs(matrix.M21), Math.Abs(matrix.M22)));
            if (largest > 0D) {
                double a = matrix.M11 / largest, b = matrix.M12 / largest, c = matrix.M21 / largest, d = matrix.M22 / largest;
                double sum = a*a+b*b+c*c+d*d, determinant = a*d-b*c;
                resolution *= largest * Math.Sqrt((sum + Math.Sqrt(Math.Max(0D, sum*sum - 4D*determinant*determinant))) / 2D);
            }
        }
        if (double.IsNaN(resolution) || double.IsInfinity(resolution) || resolution <= 0D)
            throw new InvalidDataException("XPS stroke projection resolution exceeds its numeric limit.");
        return resolution;
    }

    private static OfficeStrokeOutlineCap NativeCap(string cap) => cap == "Round" ? OfficeStrokeOutlineCap.Round :
        cap == "Square" ? OfficeStrokeOutlineCap.Square : cap == "Triangle" ? OfficeStrokeOutlineCap.Triangle : OfficeStrokeOutlineCap.Flat;
}
