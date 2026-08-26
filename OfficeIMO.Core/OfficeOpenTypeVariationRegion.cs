namespace OfficeIMO.Drawing;

/// <summary>Shared OpenType variation-region scalar semantics for metrics and CFF2 blends.</summary>
internal static class OfficeOpenTypeVariationRegion {
    internal static double CalculateScalar(double coordinate, double start, double peak, double end) {
        // OpenType 1.9.1 defines malformed or cross-zero nonzero-peak axis records as
        // non-participating. They therefore contribute a neutral multiplier of one.
        if (start > peak || peak > end || start < 0D && end > 0D && peak != 0D) return 1D;
        if (peak == 0D || coordinate == peak) return 1D;
        if (coordinate <= start || coordinate >= end) return 0D;
        if (coordinate < peak) return peak == start ? 1D : (coordinate - start) / (peak - start);
        return peak == end ? 1D : (end - coordinate) / (end - peak);
    }
}
