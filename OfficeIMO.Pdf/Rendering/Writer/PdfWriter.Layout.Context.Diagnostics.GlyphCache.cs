using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private const int DiagnosticGlyphCacheEntryLimit = 512;
        private const int DiagnosticGlyphCachePointLimit = 65536;
        // DiagnosticFont resolves these owned programs from private font snapshots;
        // caller-provided outline providers never enter this per-writer cache.
        private readonly Dictionary<(IOfficeBoundedFontProgram Font, int Glyph, double Size), List<List<OfficePoint>>>
            _diagnosticGlyphContours = new();
        private int _diagnosticGlyphPointCount;

        /// <summary>Reuses exact bounded outlines only within this writer and returns private points for placement.</summary>
        private List<List<OfficePoint>> DiagnosticGlyphOutline(IOfficeBoundedFontProgram font, int glyphId, double size) {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (font, glyphId, size);
            if (!_diagnosticGlyphContours.TryGetValue(key, out List<List<OfficePoint>>? contours)) {
                var shaped = new OfficeTextShapingResult(new[] { new OfficeShapedGlyph(glyphId, "x", 0) });
                double ascent = ((IOfficeFontBaselineMetrics)font).BaselineOffset(size);
                contours = font.GetShapedTextContoursBounded("x", shaped, 0D, -ascent, size,
                    DiagnosticGlyphCachePointLimit, cancellationToken);
                int points = 0;
                foreach (List<OfficePoint> contour in contours) points += contour.Count;
                if (_diagnosticGlyphContours.Count >= DiagnosticGlyphCacheEntryLimit
                    || points > DiagnosticGlyphCachePointLimit - _diagnosticGlyphPointCount) {
                    _diagnosticGlyphContours.Clear();
                    _diagnosticGlyphPointCount = 0;
                }
                _diagnosticGlyphContours.Add(key, contours);
                _diagnosticGlyphPointCount += points;
            }
            var result = new List<List<OfficePoint>>(contours.Count);
            foreach (List<OfficePoint> contour in contours) {
                cancellationToken.ThrowIfCancellationRequested();
                result.Add(new List<OfficePoint>(contour));
            }
            return result;
        }
    }
}
