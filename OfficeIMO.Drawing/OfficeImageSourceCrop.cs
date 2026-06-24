using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes normalized source-image crop fractions from each image edge.
/// </summary>
public readonly struct OfficeImageSourceCrop {
    /// <summary>Smallest visible source ratio used when authored crop edges collapse the source area.</summary>
    public const double MinimumVisibleRatio = 0.001D;

    /// <summary>
    /// Creates a source-image crop description from normalized edge fractions.
    /// </summary>
    public OfficeImageSourceCrop(double left, double top, double right, double bottom) {
        Left = ValidateFraction(left, nameof(left));
        Top = ValidateFraction(top, nameof(top));
        Right = ValidateFraction(right, nameof(right));
        Bottom = ValidateFraction(bottom, nameof(bottom));
    }

    /// <summary>Fraction cropped from the left edge of the source image.</summary>
    public double Left { get; }

    /// <summary>Fraction cropped from the top edge of the source image.</summary>
    public double Top { get; }

    /// <summary>Fraction cropped from the right edge of the source image.</summary>
    public double Right { get; }

    /// <summary>Fraction cropped from the bottom edge of the source image.</summary>
    public double Bottom { get; }

    /// <summary>Whether any crop edge is greater than zero.</summary>
    public bool HasCrop => Left > 0D || Top > 0D || Right > 0D || Bottom > 0D;

    /// <summary>Whether the crop leaves a non-empty source area without requiring fallback clamping.</summary>
    public bool HasVisibleSourceArea => LeavesVisibleSourceArea(Left, Top, Right, Bottom);

    /// <summary>Visible width ratio after applying left and right crop fractions.</summary>
    public double VisibleWidth => Math.Max(MinimumVisibleRatio, 1D - Left - Right);

    /// <summary>Visible height ratio after applying top and bottom crop fractions.</summary>
    public double VisibleHeight => Math.Max(MinimumVisibleRatio, 1D - Top - Bottom);

    /// <summary>
    /// Creates a source crop after clamping each input fraction into the supported normalized range.
    /// </summary>
    public static OfficeImageSourceCrop FromClampedFractions(double left, double top, double right, double bottom) =>
        new(
            ClampFraction(left),
            ClampFraction(top),
            ClampFraction(right),
            ClampFraction(bottom));

    /// <summary>
    /// Creates a source crop from normalized fractions and rejects crops that collapse the visible source area.
    /// </summary>
    public static OfficeImageSourceCrop FromStrictFractions(double left, double top, double right, double bottom) {
        var crop = new OfficeImageSourceCrop(left, top, right, bottom);
        if (!crop.HasVisibleSourceArea) {
            if (left + right >= 1D) {
                throw new ArgumentOutOfRangeException(nameof(left), "Image source crop left and right fractions must leave a visible source width.");
            }

            throw new ArgumentOutOfRangeException(nameof(top), "Image source crop top and bottom fractions must leave a visible source height.");
        }

        return crop;
    }

    /// <summary>
    /// Returns whether crop fractions leave a non-empty visible source area.
    /// </summary>
    public static bool LeavesVisibleSourceArea(double left, double top, double right, double bottom) {
        left = ValidateFraction(left, nameof(left));
        top = ValidateFraction(top, nameof(top));
        right = ValidateFraction(right, nameof(right));
        bottom = ValidateFraction(bottom, nameof(bottom));
        return left + right < 1D && top + bottom < 1D;
    }

    /// <summary>
    /// Converts the crop into a tuple.
    /// </summary>
    public (double Left, double Top, double Right, double Bottom) ToTuple() => (Left, Top, Right, Bottom);

    private static double ValidateFraction(double value, string paramName) {
        if (value < 0D || value >= 1D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Image source crop fractions must be finite values from 0 to less than 1.");
        }

        return value;
    }

    private static double ClampFraction(double value) {
        if (double.IsNaN(value) || value <= 0D) {
            return 0D;
        }

        if (double.IsInfinity(value) || value >= 1D) {
            return 1D - MinimumVisibleRatio;
        }

        return value;
    }
}
