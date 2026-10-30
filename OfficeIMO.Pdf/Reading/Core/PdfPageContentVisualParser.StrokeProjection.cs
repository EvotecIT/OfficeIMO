namespace OfficeIMO.Pdf;

internal static partial class PdfPageContentVisualParser {
    private sealed partial class Parser {
        private Action<int>? _strokeDashProjectionConsumer;
        private long _projectedStrokeDashEntries;

        internal void SetStrokeDashProjectionConsumer(Action<int>? consumer) => _strokeDashProjectionConsumer = consumer;

        private void ChargeStrokeDashProjection(int count) {
            if (_strokeDashProjectionConsumer != null) {
                _strokeDashProjectionConsumer(count);
                return;
            }
            _projectedStrokeDashEntries += count;
            if (_projectedStrokeDashEntries > _maxOperands) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperands, _maxOperands, _projectedStrokeDashEntries);
            }
        }

        private double GetRenderedStrokeWidth((double X, double Y)? direction = null) {
            double strokeWidth = _state.StrokeWidth;
            if (double.IsPositiveInfinity(strokeWidth)) return RenderedHairlineStrokeWidth;
            if (!_scaleStrokeWidthWithTransform || strokeWidth <= 0D) {
                return strokeWidth;
            }

            Matrix2D transform = _state.Transform;
            double determinant = Math.Abs((transform.A * transform.D) - (transform.B * transform.C));
            if (direction is (double dx, double dy) && determinant > 0D && (dx != 0D || dy != 0D)) {
                // A user-space pen maps to an ellipse. A single segment's painted width is the pen's
                // extent perpendicular to its page direction: w * |det| * |M^-1 v| / |v|.
                double inverseX = (transform.D * dx - transform.C * dy) / determinant;
                double inverseY = (-transform.B * dx + transform.A * dy) / determinant;
                double width = strokeWidth * determinant * Math.Sqrt(inverseX * inverseX + inverseY * inverseY) / Math.Sqrt(dx * dx + dy * dy);
                if (width > 0D && !double.IsNaN(width) && !double.IsInfinity(width)) return width;
            }
            // The area-preserving scale is exact for similarity transforms and does not inflate sheared pens.
            return determinant > 0D && !double.IsNaN(determinant) && !double.IsInfinity(determinant)
                ? strokeWidth * Math.Sqrt(determinant)
                : strokeWidth;
        }

        /// <summary>Projects dash distances and phase exactly for uniform transforms.</summary>
        private PdfStrokeDashPattern? GetRenderedStrokeDashPattern() {
            PdfStrokeDashPattern? pattern = _state.StrokeDashPattern;
            if (!pattern.HasValue || pattern.Value.Array.Count == 0) return pattern;
            // Charge the retained output before scaling or copying the array, even at identity scale.
            ChargeStrokeDashProjection(pattern.Value.Array.Count);
            double scale = GetRenderedDashScale();
            if (scale == 1D) return pattern;

            var lengths = new double[pattern.Value.Array.Count];
            for (int index = 0; index < lengths.Length; index++) lengths[index] = pattern.Value.Array[index] * scale;
            return new PdfStrokeDashPattern(lengths, pattern.Value.Phase * scale);
        }

        // Uniform transforms have an exact scalar stroke scale. Preserve the existing
        // RMS dash-distance approximation for nonuniform transforms; their diagnostic owner remains unchanged.
        private double GetRenderedDashScale() {
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
