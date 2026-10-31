using System;

namespace OfficeIMO.Drawing;

/// <summary>Shared DrawingML-compatible bounds for chart appearances.</summary>
internal static class OfficeChartStyleBounds {
    internal const double MaximumLineWidthPoints = 1584;
    internal const double HairlineWidthPoints = 0.25;

    internal static void ValidateLineWidth(double? width, string parameterName, bool allowZero = false) {
        if (width.HasValue && (double.IsNaN(width.Value) || double.IsInfinity(width.Value) ||
            (allowZero ? width.Value < 0 : width.Value <= 0) || width.Value > MaximumLineWidthPoints))
            throw new ArgumentOutOfRangeException(parameterName, "Line width must be finite and within the DrawingML limit of 1584 points.");
    }

    internal static void ValidateLineWidth(OfficeLength? width, string parameterName) {
        if (width.HasValue && (width.Value <= OfficeLength.Zero ||
            width.Value > OfficeLength.Points(MaximumLineWidthPoints)))
            throw new ArgumentOutOfRangeException(parameterName, "Line width must be greater than zero and at most 1584 points.");
    }

    internal static OfficeLength ClampNativeMarkerSize(OfficeLength size) =>
        size < OfficeLength.Points(2) ? OfficeLength.Points(2) :
        size > OfficeLength.Points(72) ? OfficeLength.Points(72) : size;
}
