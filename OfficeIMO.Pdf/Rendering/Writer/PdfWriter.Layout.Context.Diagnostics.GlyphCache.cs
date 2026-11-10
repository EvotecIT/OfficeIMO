using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private const int DiagnosticGlyphCacheEntryLimit = 512;
        private const int DiagnosticGlyphCachePointLimit = 65536;

        /// <summary>Bounded outlines of one shared glyph-paint font, reused by every document that paints with it.</summary>
        /// <remarks>
        /// Glyph-paint fonts are library-owned programs shared through their font bytes, so their outlines are pure
        /// functions of glyph and size. Entries are never mutated; placement always works on a copy.
        /// </remarks>
        private sealed class SharedDiagnosticGlyphOutlines {
            internal readonly Dictionary<(int Glyph, double Size), (List<List<OfficePoint>> Contours, DiagnosticPaintBounds Bounds)> Entries = new();
            internal int PointCount;
        }

        // Keyed weakly by the shared font program, so outlines live exactly as long as the parsed font they describe.
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<IOfficeBoundedFontProgram, SharedDiagnosticGlyphOutlines>
            SharedGlyphOutlines = new();

        /// <summary>Releases this writer's borrowed placement storage; shared outlines stay with their font program.</summary>
        private void ClearDiagnosticGlyphCache() => ClearDiagnosticGlyphWorkspace();

        private (List<List<OfficePoint>> Contours, DiagnosticPaintBounds Bounds) GetDiagnosticGlyphOutline(
            IOfficeBoundedFontProgram font, int glyphId, double size) {
            SharedDiagnosticGlyphOutlines shared = SharedGlyphOutlines.GetValue(font, static _ => new SharedDiagnosticGlyphOutlines());
            var key = (glyphId, size);
            lock (shared) {
                if (shared.Entries.TryGetValue(key, out var found)) return found;
            }
            var shaped = new OfficeTextShapingResult(new[] { new OfficeShapedGlyph(glyphId, "x", 0) });
            double ascent = ((IOfficeFontBaselineMetrics)font).BaselineOffset(size);
            List<List<OfficePoint>> contours = font.GetShapedTextContoursBounded("x", shaped, 0D, -ascent, size,
                DiagnosticGlyphCachePointLimit, cancellationToken);
            int points = 0;
            foreach (List<OfficePoint> contour in contours) points += contour.Count;
            var cached = (contours, MeasureDiagnosticPaintBounds(contours));
            lock (shared) {
                if (shared.Entries.TryGetValue(key, out var raced)) return raced;
                if (shared.Entries.Count >= DiagnosticGlyphCacheEntryLimit
                    || points > DiagnosticGlyphCachePointLimit - shared.PointCount) {
                    shared.Entries.Clear();
                    shared.PointCount = 0;
                }
                if (points <= DiagnosticGlyphCachePointLimit) {
                    shared.Entries.Add(key, cached);
                    shared.PointCount += points;
                }
            }
            return cached;
        }

        /// <summary>Reuses exact bounded outlines, borrowing placement storage only when paint still needs inspection.</summary>
        private List<List<OfficePoint>>? DiagnosticGlyphOutline(IOfficeBoundedFontProgram font, int glyphId, double size,
            GlyphPaintPlacement placement, Func<DiagnosticPaintBounds, bool>? contained) {
            cancellationToken.ThrowIfCancellationRequested();
            var cached = GetDiagnosticGlyphOutline(font, glyphId, size);
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
