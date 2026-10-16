namespace OfficeIMO.Pdf;

internal static partial class PdfPageContentVisualParser {
    private sealed partial class Parser {
        private double GetRenderedStrokeWidth() {
            double width = _state.StrokeWidth;
            if (double.IsPositiveInfinity(width)) return RenderedHairlineStrokeWidth;
            return width <= 0D ? width : width * GetRenderedStrokeScale();
        }

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

        // Uniform transforms have an exact scalar stroke scale. Preserve the existing
        // RMS approximation for nonuniform transforms; their diagnostic owner remains unchanged.
        private double GetRenderedStrokeScale() {
            if (!_scaleStrokeWidthWithTransform) return 1D;
            Matrix2D transform = _state.Transform;
            double squaredScale = ((transform.A * transform.A) +
                                   (transform.B * transform.B) +
                                   (transform.C * transform.C) +
                                   (transform.D * transform.D)) / 2D;
            return squaredScale > 0D && !double.IsNaN(squaredScale) && !double.IsInfinity(squaredScale)
                ? Math.Sqrt(squaredScale)
                : 1D;
        }
    }
}
