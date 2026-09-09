namespace OfficeIMO.Pdf;

/// <summary>Internal quarter-turn reading coordinates; public artifacts retain source-page geometry.</summary>
internal sealed partial class PdfUnderstandingReadingFrame {
    private readonly PdfUnderstandingPageContext _context;
    private readonly double _cos;
    private readonly double _sin;
    private readonly double _left;
    private readonly double _bottom;
    private readonly Dictionary<PdfTextSpan, PdfTextSpan> _originalRuns = new();
    private readonly Dictionary<PdfUnderstandingWord, PdfUnderstandingWord> _originalWords = new();
    private readonly Dictionary<PdfUnderstandingWord, PdfUnderstandingWord> _projectedWords = new();
    private readonly Dictionary<PdfTextSpan, PdfTextSpan> _projectedRuns = new();

    private PdfUnderstandingReadingFrame(PdfUnderstandingPageContext context, int angle) {
        _context = context;
        Angle = angle;
        _cos = angle == 0 ? 1D : angle == 180 ? -1D : 0D;
        _sin = angle == 90 ? 1D : angle == 270 ? -1D : 0D;
        PdfPageBox? box = context.Page.GetGeometry().EffectiveBox;
        double left = box?.Left ?? 0D, bottom = box?.Bottom ?? 0D;
        double right = left + context.Width, top = bottom + context.Height;
        var corners = new[] { Rotate(left, bottom), Rotate(left, top), Rotate(right, bottom), Rotate(right, top) };
        _left = corners.Min(static point => point.X);
        _bottom = corners.Min(static point => point.Y);
        Width = corners.Max(static point => point.X) - _left;
        Height = corners.Max(static point => point.Y) - _bottom;
    }

    internal int Angle { get; }
    internal double Width { get; }
    internal double Height { get; }

    internal static PdfUnderstandingReadingFrame? TryCreate(PdfUnderstandingPageContext context, IReadOnlyList<PdfTextSpan> runs) {
        return TryCreate(context, runs, static run => run.Text, static run => run.RotationDegrees, false);
    }

    internal static PdfUnderstandingReadingFrame CreateForPositionedWords(PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingWord> words) =>
        TryCreate(context, words, static word => word.Text, static word => word.RotationDegrees, true)!;

    private static PdfUnderstandingReadingFrame? TryCreate<T>(PdfUnderstandingPageContext context,
        IReadOnlyList<T> runs, Func<T, string> text, Func<T, double> rotation, bool includeIdentity) {
        var weights = new long[4];
        long total = 0;
        foreach (T run in runs) {
            context.ConsumeWork();
            string value = text(run);
            if (string.IsNullOrWhiteSpace(value)) continue;
            int weight = value.Length;
            total += weight;
            double angle = (rotation(run) % 360D + 360D) % 360D;
            if (double.IsNaN(angle) || double.IsInfinity(angle)) continue;
            int quarter = (int)Math.Round(angle / 90D) % 4;
            double distance = Math.Abs(angle - quarter * 90D);
            if (Math.Min(distance, 360D - distance) <= 2D) weights[quarter] += weight;
        }
        if (total == 0) return includeIdentity ? new PdfUnderstandingReadingFrame(context, 0) : null;
        int dominant = 0;
        for (int quarter = 1; quarter < 4; quarter++) if (weights[quarter] > weights[dominant]) dominant = quarter;
        return dominant != 0 && weights[dominant] >= total * 0.7D
            ? new PdfUnderstandingReadingFrame(context, dominant * 90)
            : includeIdentity ? new PdfUnderstandingReadingFrame(context, 0) : null;
    }

    private (double X, double Y) Rotate(double x, double y) => (_cos * x + _sin * y, -_sin * x + _cos * y);
    internal (double X, double Y) ToFrame(double x, double y) {
        (double px, double py) = Rotate(x, y);
        return (px - _left, py - _bottom);
    }
    private (double X, double Y) ToSource(double x, double y) {
        x += _left; y += _bottom;
        return (_cos * x - _sin * y, _sin * x + _cos * y);
    }

    internal IReadOnlyList<PdfTextSpan> ProjectRuns(IReadOnlyList<PdfTextSpan> source) {
        var result = new PdfTextSpan[source.Count];
        for (int index = 0; index < result.Length; index++) {
            _context.ConsumeWork();
            PdfTextSpan run = source[index];
            (double x, double y) = ToFrame(run.X, run.Y);
            PdfTextSpan projected = run.WithLayoutGeometry(x, y, PdfAdvancedUnderstandingStages.NormalizeAngle(run.RotationDegrees - Angle), _context.Height);
            result[index] = projected;
            _originalRuns.Add(projected, run);
            _projectedRuns.Add(run, projected);
        }
        return Array.AsReadOnly(result);
    }

    internal PdfVisualBounds ImageBounds(PdfVisualBounds sourceVisual) {
        PdfVisualBounds user = _context.Page.TransformVisualBoundsToUser(sourceVisual.Left, sourceVisual.Top, sourceVisual.Right, sourceVisual.Bottom);
        var points = new[] { ToFrame(user.Left, user.Top), ToFrame(user.Left, user.Bottom), ToFrame(user.Right, user.Top), ToFrame(user.Right, user.Bottom) };
        return new PdfVisualBounds(points.Min(static point => point.X), Height - points.Max(static point => point.Y),
            points.Max(static point => point.X), Height - points.Min(static point => point.Y));
    }

    private PdfLogicalVisualBounds SourceVisualBounds(double left, double bottom, double right, double top) {
        var points = new[] { ToSource(left, bottom), ToSource(left, top), ToSource(right, bottom), ToSource(right, top) };
        PdfVisualBounds visual = _context.Page.TransformBoundsToVisual(points.Min(static point => point.X), points.Min(static point => point.Y),
            points.Max(static point => point.X), points.Max(static point => point.Y));
        return new PdfLogicalVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom);
    }
}
