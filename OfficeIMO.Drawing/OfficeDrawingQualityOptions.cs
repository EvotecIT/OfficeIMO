using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Options for shared drawing quality checks.
/// </summary>
public sealed class OfficeDrawingQualityOptions {
    /// <summary>Default drawing quality options.</summary>
    public static OfficeDrawingQualityOptions Default { get; } = new OfficeDrawingQualityOptions();

    /// <summary>
    /// Creates drawing quality options.
    /// </summary>
    /// <param name="boundsTolerance">Allowed canvas-bound overflow in drawing units.</param>
    /// <param name="overlapTolerance">Allowed text-box overlap in drawing units.</param>
    /// <param name="detectTextOverlap">Whether text-box overlap diagnostics are emitted.</param>
    public OfficeDrawingQualityOptions(double boundsTolerance = 0.01D, double overlapTolerance = 0.01D, bool detectTextOverlap = true) {
        BoundsTolerance = ValidateFiniteNonNegative(boundsTolerance, nameof(boundsTolerance));
        OverlapTolerance = ValidateFiniteNonNegative(overlapTolerance, nameof(overlapTolerance));
        DetectTextOverlap = detectTextOverlap;
    }

    /// <summary>Allowed canvas-bound overflow in drawing units.</summary>
    public double BoundsTolerance { get; }

    /// <summary>Allowed text-box overlap in drawing units.</summary>
    public double OverlapTolerance { get; }

    /// <summary>Whether text-box overlap diagnostics are emitted.</summary>
    public bool DetectTextOverlap { get; }

    private static double ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing quality tolerances must be finite non-negative numbers.");
        }

        return value;
    }
}
