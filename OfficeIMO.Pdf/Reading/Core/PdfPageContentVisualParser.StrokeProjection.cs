namespace OfficeIMO.Pdf;

internal static partial class PdfPageContentVisualParser {
    private sealed partial class Parser {
        /// <summary>Projects exact dash distances and phase with the same transform as stroke width.</summary>
        private PdfStrokeDashPattern? GetRenderedStrokeDashPattern() {
            PdfStrokeDashPattern? pattern = _state.StrokeDashPattern;
            if (!pattern.HasValue || pattern.Value.Array.Count == 0) return pattern;
            double scale = GetRenderedStrokeScale();
            if (scale == 1D) return pattern;

            var lengths = new double[pattern.Value.Array.Count];
            for (int index = 0; index < lengths.Length; index++) lengths[index] = pattern.Value.Array[index] * scale;
            return new PdfStrokeDashPattern(lengths, pattern.Value.Phase * scale);
        }

        // Use the shared PDF-text candidate's area-preserving approximation for general
        // paths. Individual line widths retain its perpendicular-pen projection.
        private double GetRenderedStrokeScale() {
            if (!_scaleStrokeWidthWithTransform) return 1D;
            Matrix2D transform = _state.Transform;
            double determinant = Math.Abs((transform.A * transform.D) - (transform.B * transform.C));
            return determinant > 0D && !double.IsNaN(determinant) && !double.IsInfinity(determinant)
                ? Math.Sqrt(determinant)
                : 1D;
        }
    }
}
