using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Uses the selected PDF glyphs and PDF-unit positioning, never a separately shaped text approximation.</summary>
        private IEnumerable<(List<List<OfficePoint>> Contours, bool NominalBounds)> DiagnosticGlyphContours(RichSeg segment, string text, double x, double baseline) {
            if (string.IsNullOrWhiteSpace(text)) yield break;
            foreach (var glyph in EnumerateGlyphPaintContours(segment, text, x, baseline, currentOpts,
                _diagnosticFonts, DiagnosticTextOptions, cancellationToken: cancellationToken,
                glyphOutline: DiagnosticGlyphOutline, fontCacheCleared: ClearDiagnosticGlyphCache)) yield return glyph;
        }

        private IEnumerable<List<List<OfficePoint>>> DiagnosticDecorationContours(double left, double right, double y, OfficeTextDecorationStyle style) {
            const double thickness = .5D;
            if (style == OfficeTextDecorationStyle.Double) {
                double separation = Math.Max(thickness * 1.8D, .8D);
                foreach (var part in DiagnosticDecorationContours(left, right, y - separation / 2D, OfficeTextDecorationStyle.Single)) yield return part;
                foreach (var part in DiagnosticDecorationContours(left, right, y + separation / 2D, OfficeTextDecorationStyle.Single)) yield return part;
                yield break;
            }
            var points = new List<OfficePoint> { new(left, y) };
            if (style == OfficeTextDecorationStyle.Wavy) {
                double step = Math.Max(thickness * 5D, 2D), amplitude = Math.Max(thickness * 1.5D, .75D);
                bool raised = true;
                for (double cursor = left; cursor < right;) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (points.Count >= 65536 || cursor + step <= cursor) throw new NotSupportedException("Decoration paint exceeds its bounded contour inspection limit.");
                    cursor = Math.Min(right, cursor + step);
                    points.Add(new OfficePoint(cursor, y + (raised ? amplitude : -amplitude))); raised = !raised;
                }
            } else points.Add(new OfficePoint(right, y));
            double[]? dash = style == OfficeTextDecorationStyle.Dashed ? new[] { 2.5D, 1.5D }
                : style == OfficeTextDecorationStyle.Dotted ? new[] { .5D, 1.5D } : null;
            var contours = new[] { new OfficeFlattenedPathContour(points, false) };
            yield return OfficeStrokeGeometry.Create(contours, thickness,
                style == OfficeTextDecorationStyle.Dotted ? OfficeStrokeLineCap.Round : OfficeStrokeLineCap.Butt,
                style == OfficeTextDecorationStyle.Wavy ? OfficeStrokeLineJoin.Round : OfficeStrokeLineJoin.Miter,
                10D, dash, 0D, 4D, double.NegativeInfinity, double.NegativeInfinity, double.PositiveInfinity, double.PositiveInfinity, cancellationToken);
        }
    }
}
