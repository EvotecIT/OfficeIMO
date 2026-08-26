using System;

namespace OfficeIMO.Drawing;

/// <summary>Shared OpenType variation-region scalar semantics for metrics and CFF2 blends.</summary>
internal static class OfficeOpenTypeVariationRegion {
    internal static double CalculateTupleScalar(
        double coordinate,
        double peak,
        double? intermediateStart,
        double? intermediateEnd) {
        if (intermediateStart.HasValue && intermediateEnd.HasValue) {
            double start = intermediateStart.Value;
            double end = intermediateEnd.Value;
            // Invalid or cross-zero nonzero-peak records do not participate in this
            // tuple. They must not suppress variation contributed by another axis.
            if (start > peak || peak > end || start < 0D && end > 0D && peak != 0D) return 1D;
            // A zero peak means this axis does not participate in the tuple. Its
            // intermediate bounds must not suppress variation driven by another axis.
            if (peak == 0D) return 1D;
            if (coordinate < start || coordinate > end) return 0D;
            if (coordinate < peak) return peak == start ? 1D : (coordinate - start) / (peak - start);
            if (coordinate > peak) return peak == end ? 1D : (end - coordinate) / (end - peak);
            return 1D;
        }

        if (peak == 0D) return 1D;
        // OpenType defines a non-intermediate region's implicit bounds as the zero origin
        // and the peak tuple. A same-sign coordinate beyond that peak is outside the region.
        if (coordinate == 0D || coordinate < 0D != peak < 0D || Math.Abs(coordinate) > Math.Abs(peak)) return 0D;
        return coordinate / peak;
    }

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
