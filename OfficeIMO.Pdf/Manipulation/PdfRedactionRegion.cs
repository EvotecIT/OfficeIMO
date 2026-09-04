namespace OfficeIMO.Pdf;

/// <summary>Geometry kind supplied by a redaction-review surface.</summary>
public enum PdfRedactionRegionKind {
    /// <summary>One axis-aligned rectangle in PDF user space.</summary>
    Rectangle,
    /// <summary>A four-point region represented conservatively for destructive application.</summary>
    Quadrilateral,
    /// <summary>A polygon represented conservatively for destructive application.</summary>
    Polygon,
    /// <summary>A stroked freehand path represented by conservative segment bounds.</summary>
    Freehand,
    /// <summary>A caller-defined group of already normalized redaction rectangles.</summary>
    Group
}

/// <summary>One point in PDF default user-space coordinates.</summary>
public readonly struct PdfRedactionPoint {
    /// <summary>Creates a finite PDF point.</summary>
    public PdfRedactionPoint(double x, double y) {
        if (double.IsNaN(x) || double.IsInfinity(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (double.IsNaN(y) || double.IsInfinity(y)) throw new ArgumentOutOfRangeException(nameof(y));
        X = x;
        Y = y;
    }

    /// <summary>Horizontal coordinate in PDF points.</summary>
    public double X { get; }
    /// <summary>Vertical coordinate in PDF points.</summary>
    public double Y { get; }
}

/// <summary>
/// Review geometry normalized into one or more conservative rectangles used consistently by planning,
/// destructive application, and evidence. Conservative normalization may redact more than the visual path,
/// but never less than its declared bounds.
/// </summary>
public sealed class PdfRedactionRegion {
    private PdfRedactionRegion(PdfRedactionRegionKind kind, int pageNumber, IReadOnlyList<PdfRedactionPoint> points, IReadOnlyList<PdfRedactionArea> areas, string? label, double strokeWidth) {
        Kind = kind;
        PageNumber = pageNumber;
        Points = points;
        Areas = areas;
        Label = label;
        StrokeWidth = strokeWidth;
    }

    /// <summary>Original review geometry kind.</summary>
    public PdfRedactionRegionKind Kind { get; }
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Original points, when the region is path based.</summary>
    public IReadOnlyList<PdfRedactionPoint> Points { get; }
    /// <summary>Canonical destructive rectangles.</summary>
    public IReadOnlyList<PdfRedactionArea> Areas { get; }
    /// <summary>Optional caller label.</summary>
    public string? Label { get; }
    /// <summary>Freehand stroke width in PDF points, or zero for other kinds.</summary>
    public double StrokeWidth { get; }

    /// <summary>Creates one rectangle region.</summary>
    public static PdfRedactionRegion Rectangle(int pageNumber, double x, double y, double width, double height, string? label = null) {
        var area = new PdfRedactionArea(pageNumber, x, y, width, height, label);
        return new PdfRedactionRegion(PdfRedactionRegionKind.Rectangle, pageNumber, Array.Empty<PdfRedactionPoint>(), new[] { area }, label, 0D);
    }

    /// <summary>Creates a quadrilateral whose complete axis-aligned bounds are removed.</summary>
    public static PdfRedactionRegion Quadrilateral(int pageNumber, IEnumerable<PdfRedactionPoint> points, string? label = null) =>
        FromBoundedPoints(PdfRedactionRegionKind.Quadrilateral, pageNumber, points, 4, label);

    /// <summary>Creates a polygon whose complete axis-aligned bounds are removed.</summary>
    public static PdfRedactionRegion Polygon(int pageNumber, IEnumerable<PdfRedactionPoint> points, string? label = null) =>
        FromBoundedPoints(PdfRedactionRegionKind.Polygon, pageNumber, points, 3, label);

    /// <summary>Creates a freehand region as conservative per-segment rectangles.</summary>
    public static PdfRedactionRegion Freehand(int pageNumber, IEnumerable<PdfRedactionPoint> points, double strokeWidth, string? label = null) {
        ValidatePage(pageNumber);
        if (double.IsNaN(strokeWidth) || double.IsInfinity(strokeWidth) || strokeWidth <= 0D) throw new ArgumentOutOfRangeException(nameof(strokeWidth));
        PdfRedactionPoint[] path = SnapshotPoints(points, 2);
        double radius = strokeWidth / 2D;
        var areas = new PdfRedactionArea[path.Length - 1];
        for (int index = 1; index < path.Length; index++) {
            PdfRedactionPoint left = path[index - 1];
            PdfRedactionPoint right = path[index];
            double x = Math.Min(left.X, right.X) - radius;
            double y = Math.Min(left.Y, right.Y) - radius;
            double width = Math.Max(strokeWidth, Math.Abs(left.X - right.X) + strokeWidth);
            double height = Math.Max(strokeWidth, Math.Abs(left.Y - right.Y) + strokeWidth);
            areas[index - 1] = new PdfRedactionArea(pageNumber, x, y, width, height, label);
        }
        return new PdfRedactionRegion(PdfRedactionRegionKind.Freehand, pageNumber, path, areas, label, strokeWidth);
    }

    /// <summary>Groups pre-normalized rectangles from one page into one review decision.</summary>
    public static PdfRedactionRegion Group(int pageNumber, IEnumerable<PdfRedactionArea> areas, string? label = null) {
        ValidatePage(pageNumber);
        Guard.NotNull(areas, nameof(areas));
        PdfRedactionArea[] snapshot = areas.ToArray();
        if (snapshot.Length == 0) throw new ArgumentException("A region group requires at least one area.", nameof(areas));
        if (snapshot.Any(area => area.PageNumber != pageNumber)) throw new ArgumentException("Every grouped area must use the region page number.", nameof(areas));
        return new PdfRedactionRegion(PdfRedactionRegionKind.Group, pageNumber, Array.Empty<PdfRedactionPoint>(), snapshot, label, 0D);
    }

    /// <summary>Creates one atomic review region from a standard PDF /Redact annotation, honoring every valid /QuadPoints quadrilateral before falling back to /Rect.</summary>
    public static PdfRedactionRegion FromRedactAnnotation(PdfAnnotation annotation) {
        Guard.NotNull(annotation, nameof(annotation));
        if (!string.Equals(annotation.Subtype, "Redact", StringComparison.Ordinal)) throw new ArgumentException("The annotation subtype must be Redact.", nameof(annotation));
        if (!annotation.PageNumber.HasValue) throw new ArgumentException("The redaction annotation must identify its page.", nameof(annotation));
        int pageNumber = annotation.PageNumber.Value;
        if (annotation.QuadPoints.Count >= 8 && annotation.QuadPoints.Count % 8 == 0) {
            var areas = new List<PdfRedactionArea>(annotation.QuadPoints.Count / 8);
            for (int offset = 0; offset < annotation.QuadPoints.Count; offset += 8) {
                var points = new PdfRedactionPoint[4];
                bool valid = true;
                for (int point = 0; point < 4; point++) {
                    double x = annotation.QuadPoints[offset + point * 2];
                    double y = annotation.QuadPoints[offset + point * 2 + 1];
                    if (double.IsNaN(x) || double.IsInfinity(x) || double.IsNaN(y) || double.IsInfinity(y)) { valid = false; break; }
                    points[point] = new PdfRedactionPoint(x, y);
                }
                if (!valid) { areas.Clear(); break; }
                try { areas.Add(Quadrilateral(pageNumber, points, annotation.Name).Areas[0]); }
                catch (ArgumentException) { areas.Clear(); break; }
            }
            if (areas.Count > 0) return Group(pageNumber, areas, annotation.Name);
        }
        if (annotation.Width <= 0D || annotation.Height <= 0D) throw new ArgumentException("The redaction annotation requires valid /QuadPoints or a positive /Rect.", nameof(annotation));
        return Rectangle(pageNumber, annotation.X1, annotation.Y1, annotation.Width, annotation.Height, annotation.Name);
    }

    private static PdfRedactionRegion FromBoundedPoints(PdfRedactionRegionKind kind, int pageNumber, IEnumerable<PdfRedactionPoint> points, int requiredCount, string? label) {
        ValidatePage(pageNumber);
        PdfRedactionPoint[] snapshot = SnapshotPoints(points, requiredCount);
        if (kind == PdfRedactionRegionKind.Quadrilateral && snapshot.Length != 4) throw new ArgumentException("A quadrilateral requires exactly four points.", nameof(points));
        double left = snapshot.Min(static point => point.X);
        double right = snapshot.Max(static point => point.X);
        double bottom = snapshot.Min(static point => point.Y);
        double top = snapshot.Max(static point => point.Y);
        if (right <= left || top <= bottom) throw new ArgumentException("Region points must enclose a positive area.", nameof(points));
        return new PdfRedactionRegion(kind, pageNumber, snapshot, new[] { new PdfRedactionArea(pageNumber, left, bottom, right - left, top - bottom, label) }, label, 0D);
    }

    private static PdfRedactionPoint[] SnapshotPoints(IEnumerable<PdfRedactionPoint> points, int minimumCount) {
        Guard.NotNull(points, nameof(points));
        PdfRedactionPoint[] snapshot = points.ToArray();
        if (snapshot.Length < minimumCount) throw new ArgumentException($"The region requires at least {minimumCount} points.", nameof(points));
        return snapshot;
    }

    private static void ValidatePage(int pageNumber) {
        Guard.PositiveInteger(pageNumber, nameof(pageNumber));
    }
}
