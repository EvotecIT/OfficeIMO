using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly List<List<OfficePoint>> _diagnosticGlyphBuffers = new();
        private readonly List<List<OfficePoint>> _diagnosticPlacedGlyph = new();
        private int _diagnosticGlyphBufferCapacity;

        /// <summary>Releases the bounded placement workspace together with its origin-outline cache.</summary>
        private void ClearDiagnosticGlyphWorkspace() {
            _diagnosticGlyphBuffers.Clear();
            _diagnosticPlacedGlyph.Clear();
            _diagnosticGlyphBufferCapacity = 0;
        }

        /// <summary>Copies immutable origin points before the ordinary placement and paint transforms mutate them.</summary>
        /// <remarks>The result is borrowed until the next glyph. Diagnostic paint consumes it synchronously and never retains it.</remarks>
        private List<List<OfficePoint>> CopyDiagnosticGlyphForPlacement(List<List<OfficePoint>> contours) {
            int requiredCapacity = 0;
            for (int i = 0; i < contours.Count; i++) requiredCapacity += contours[i].Count;
            if (contours.Count > DiagnosticGlyphCacheEntryLimit || requiredCapacity > DiagnosticGlyphCachePointLimit) {
                var owned = new List<List<OfficePoint>>(contours.Count);
                foreach (List<OfficePoint> contour in contours) {
                    cancellationToken.ThrowIfCancellationRequested();
                    owned.Add(new List<OfficePoint>(contour));
                }
                return owned;
            }

            int growth = 0;
            for (int i = 0; i < contours.Count; i++) {
                int capacity = i < _diagnosticGlyphBuffers.Count ? _diagnosticGlyphBuffers[i].Capacity : 0;
                growth += Math.Max(0, contours[i].Count - capacity);
            }
            // Include unused buffers: alternating contour shapes must not grow
            // retained storage beyond the same point budget as the origin cache.
            if (growth > DiagnosticGlyphCachePointLimit - _diagnosticGlyphBufferCapacity) ClearDiagnosticGlyphWorkspace();
            _diagnosticPlacedGlyph.Clear();
            for (int i = 0; i < contours.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                List<OfficePoint> origin = contours[i];
                if (i == _diagnosticGlyphBuffers.Count) _diagnosticGlyphBuffers.Add(new List<OfficePoint>());
                List<OfficePoint> target = _diagnosticGlyphBuffers[i];
                if (target.Capacity < origin.Count) {
                    _diagnosticGlyphBufferCapacity += origin.Count - target.Capacity;
                    target.Capacity = origin.Count;
                }
                target.Clear();
                target.AddRange(origin);
                _diagnosticPlacedGlyph.Add(target);
            }
            return _diagnosticPlacedGlyph;
        }
    }
}
