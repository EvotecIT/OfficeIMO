using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private const int DiagnosticGlyphCacheEntryLimit = 512;
        private const int DiagnosticGlyphCachePointLimit = 65536;
        // Shared glyph-paint font resolution uses owned private font snapshots;
        // caller-provided outline providers never enter this per-writer cache.
        private readonly Dictionary<(IOfficeBoundedFontProgram Font, int Glyph, double Size),
            (List<List<OfficePoint>> Contours, DiagnosticPaintBounds Bounds)>
            _diagnosticGlyphContours = new();
        private int _diagnosticGlyphPointCount;

        /// <summary>Releases retained font programs with their cached outlines.</summary>
        private void ClearDiagnosticGlyphCache() {
            _diagnosticGlyphContours.Clear();
            _diagnosticGlyphPointCount = 0;
            ClearDiagnosticGlyphWorkspace();
        }

        /// <summary>Reuses exact bounded outlines, borrowing placement storage only when paint still needs inspection.</summary>
        private List<List<OfficePoint>>? DiagnosticGlyphOutline(IOfficeBoundedFontProgram font, int glyphId, double size,
            GlyphPaintPlacement placement, Func<DiagnosticPaintBounds, bool>? contained) {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (font, glyphId, size);
            if (!_diagnosticGlyphContours.TryGetValue(key, out var cached)) {
                var shaped = new OfficeTextShapingResult(new[] { new OfficeShapedGlyph(glyphId, "x", 0) });
                double ascent = ((IOfficeFontBaselineMetrics)font).BaselineOffset(size);
                List<List<OfficePoint>> contours = font.GetShapedTextContoursBounded("x", shaped, 0D, -ascent, size,
                    DiagnosticGlyphCachePointLimit, cancellationToken);
                int points = 0;
                foreach (List<OfficePoint> contour in contours) points += contour.Count;
                if (_diagnosticGlyphContours.Count >= DiagnosticGlyphCacheEntryLimit
                    || points > DiagnosticGlyphCachePointLimit - _diagnosticGlyphPointCount) {
                    ClearDiagnosticGlyphCache();
                }
                cached = (contours, MeasureDiagnosticPaintBounds(contours));
                _diagnosticGlyphContours.Add(key, cached);
                _diagnosticGlyphPointCount += points;
            }
            // A glyph without outline points paints no ink, so no clip can remove any. Its inspection would
            // measure only the clips; they also bound every inked glyph of the run, which reports them.
            if (contained != null && HasNoOutlinePoints(cached.Contours)) return null;
            // These are actual bounded contour extrema, never an unverified glyf header box.
            // Positive unsheared placement is monotone on each axis, including the Y reflection.
            if (contained != null && font is OfficeTrueTypeFont && cached.Bounds.IsFinite
                && placement.Shear == 0D && !placement.SyntheticBold
                && placement.HorizontalScale > 0D && !double.IsInfinity(placement.HorizontalScale)) {
                OfficePoint first = placement.TransformPoint(new OfficePoint(cached.Bounds.Left, cached.Bounds.Bottom));
                OfficePoint second = placement.TransformPoint(new OfficePoint(cached.Bounds.Right, cached.Bounds.Top));
                var bounds = new DiagnosticPaintBounds(first.X, first.Y, second.X, second.Y);
                if (bounds.IsFinite && contained(bounds)) return null;
            }
            return CopyDiagnosticGlyphForPlacement(cached.Contours);
        }

        private static bool HasNoOutlinePoints(List<List<OfficePoint>> contours) {
            foreach (List<OfficePoint> contour in contours) if (contour.Count > 0) return false;
            return true;
        }
    }
}
