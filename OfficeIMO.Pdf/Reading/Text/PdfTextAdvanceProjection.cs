namespace OfficeIMO.Pdf;

/// <summary>Projects signed text-space character advances onto a span's resolved baseline direction.</summary>
internal static class PdfTextAdvanceProjection {
    internal static bool TryGetResolvedBoundaries(PdfTextSpan span, out double[] boundaries) {
        IReadOnlyList<double>? advances = span.CharacterAdvances;
        if (advances is null || advances.Count != span.Text.Length) {
            boundaries = Array.Empty<double>();
            return false;
        }

        double signedTotal = 0D;
        for (int i = 0; i < advances.Count; i++) {
            double advance = advances[i];
            if (!IsFinite(advance)) {
                boundaries = Array.Empty<double>();
                return false;
            }
            signedTotal += advance;
            if (!IsFinite(signedTotal)) {
                boundaries = Array.Empty<double>();
                return false;
            }
        }
        if (Math.Abs(signedTotal) <= double.Epsilon) {
            boundaries = Array.Empty<double>();
            return false;
        }

        // Parser-produced spans bind RotationDegrees to the painted baseline, independently
        // of spacing-inclusive movement. Older synthetic spans do not carry that direction,
        // so retain aggregate-direction normalization as their compatibility fallback.
        double directionSign = span.CharacterAdvanceDirection != 0D
            ? span.CharacterAdvanceDirection
            : signedTotal < 0D ? -1D : 1D;
        boundaries = new double[advances.Count + 1];
        for (int i = 0; i < advances.Count; i++) {
            boundaries[i + 1] = boundaries[i] + advances[i] * directionSign;
            if (!IsFinite(boundaries[i + 1])) {
                boundaries = Array.Empty<double>();
                return false;
            }
        }
        return true;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
