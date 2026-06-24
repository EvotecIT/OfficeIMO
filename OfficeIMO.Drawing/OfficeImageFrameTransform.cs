using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes rotation and mirror transforms applied to an image frame in destination coordinates.
/// </summary>
public readonly struct OfficeImageFrameTransform : IEquatable<OfficeImageFrameTransform> {
    private const double RotationEpsilon = 0.000001D;

    /// <summary>
    /// Creates a destination-space image frame transform.
    /// </summary>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="centerX">Rotation and mirror center X coordinate.</param>
    /// <param name="centerY">Rotation and mirror center Y coordinate.</param>
    /// <param name="flipHorizontal">Whether to mirror horizontally around the center.</param>
    /// <param name="flipVertical">Whether to mirror vertically around the center.</param>
    public OfficeImageFrameTransform(double rotationDegrees, double centerX, double centerY, bool flipHorizontal = false, bool flipVertical = false) {
        ValidateFinite(rotationDegrees, nameof(rotationDegrees));
        ValidateFinite(centerX, nameof(centerX));
        ValidateFinite(centerY, nameof(centerY));

        RotationDegrees = rotationDegrees;
        CenterX = centerX;
        CenterY = centerY;
        FlipHorizontal = flipHorizontal;
        FlipVertical = flipVertical;
    }

    /// <summary>Clockwise rotation in degrees.</summary>
    public double RotationDegrees { get; }

    /// <summary>Rotation and mirror center X coordinate.</summary>
    public double CenterX { get; }

    /// <summary>Rotation and mirror center Y coordinate.</summary>
    public double CenterY { get; }

    /// <summary>Whether to mirror horizontally around the center.</summary>
    public bool FlipHorizontal { get; }

    /// <summary>Whether to mirror vertically around the center.</summary>
    public bool FlipVertical { get; }

    /// <summary>Horizontal frame scale used for mirror transforms.</summary>
    public double ScaleX => FlipHorizontal ? -1D : 1D;

    /// <summary>Vertical frame scale used for mirror transforms.</summary>
    public double ScaleY => FlipVertical ? -1D : 1D;

    /// <summary>Whether the frame applies rotation.</summary>
    public bool HasRotation => Math.Abs(RotationDegrees) >= RotationEpsilon;

    /// <summary>Whether the frame applies any mirror transform.</summary>
    public bool HasFlip => FlipHorizontal || FlipVertical;

    /// <summary>Whether the frame applies rotation or mirroring.</summary>
    public bool HasTransform => HasRotation || HasFlip;

    /// <summary>
    /// Creates an affine matrix that applies this frame transform to destination-space coordinates.
    /// </summary>
    /// <remarks>
    /// Mirroring is applied before rotation around the shared center, matching Office image behavior
    /// and the SVG transform sequence used by <see cref="OfficeSvgImageRenderer" />.
    /// </remarks>
    public OfficeTransform CreateDestinationTransform() {
        if (!HasTransform) {
            return OfficeTransform.Identity;
        }

        return OfficeTransform.Translate(-CenterX, -CenterY)
            .Then(OfficeTransform.Scale(ScaleX, ScaleY))
            .Then(OfficeTransform.RotateDegrees(RotationDegrees))
            .Then(OfficeTransform.Translate(CenterX, CenterY));
    }

    /// <summary>
    /// Converts the frame transform into a tuple useful for assertions and diagnostics.
    /// </summary>
    public (double RotationDegrees, double CenterX, double CenterY, bool FlipHorizontal, bool FlipVertical) ToTuple() =>
        (RotationDegrees, CenterX, CenterY, FlipHorizontal, FlipVertical);

    /// <inheritdoc />
    public bool Equals(OfficeImageFrameTransform other) =>
        RotationDegrees.Equals(other.RotationDegrees) &&
        CenterX.Equals(other.CenterX) &&
        CenterY.Equals(other.CenterY) &&
        FlipHorizontal == other.FlipHorizontal &&
        FlipVertical == other.FlipVertical;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeImageFrameTransform other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = RotationDegrees.GetHashCode();
            hash = (hash * 397) ^ CenterX.GetHashCode();
            hash = (hash * 397) ^ CenterY.GetHashCode();
            hash = (hash * 397) ^ FlipHorizontal.GetHashCode();
            hash = (hash * 397) ^ FlipVertical.GetHashCode();
            return hash;
        }
    }

    /// <summary>Returns true when two frame transforms are equal.</summary>
    public static bool operator ==(OfficeImageFrameTransform left, OfficeImageFrameTransform right) => left.Equals(right);

    /// <summary>Returns true when two frame transforms are not equal.</summary>
    public static bool operator !=(OfficeImageFrameTransform left, OfficeImageFrameTransform right) => !left.Equals(right);

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Image frame transform values must be finite.");
        }
    }
}
