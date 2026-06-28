using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes a projected image render operation shared by raster and SVG renderers.
/// </summary>
public readonly struct OfficeImageProjection {
    /// <summary>
    /// Creates an image projection from destination placement, source crop, and transform settings.
    /// </summary>
    /// <param name="placement">Destination placement rectangle.</param>
    /// <param name="sourceCrop">Normalized source crop fractions.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Optional rotation center X coordinate. Defaults to destination center.</param>
    /// <param name="rotationCenterY">Optional rotation center Y coordinate. Defaults to destination center.</param>
    /// <param name="flipHorizontal">Whether to mirror horizontally around the rotation center.</param>
    /// <param name="flipVertical">Whether to mirror vertically around the rotation center.</param>
    public OfficeImageProjection(
        OfficeImagePlacement placement,
        OfficeImageSourceCrop sourceCrop = default,
        double rotationDegrees = 0D,
        double? rotationCenterX = null,
        double? rotationCenterY = null,
        bool flipHorizontal = false,
        bool flipVertical = false) {
        EnsureFinite(rotationDegrees, nameof(rotationDegrees));

        Placement = placement;
        SourceCrop = sourceCrop;
        RotationDegrees = rotationDegrees;
        RotationCenterX = rotationCenterX ?? placement.X + (placement.Width / 2D);
        RotationCenterY = rotationCenterY ?? placement.Y + (placement.Height / 2D);
        EnsureFinite(RotationCenterX, nameof(rotationCenterX));
        EnsureFinite(RotationCenterY, nameof(rotationCenterY));
        FlipHorizontal = flipHorizontal;
        FlipVertical = flipVertical;
    }

    /// <summary>Destination placement rectangle.</summary>
    public OfficeImagePlacement Placement { get; }

    /// <summary>Normalized source crop fractions.</summary>
    public OfficeImageSourceCrop SourceCrop { get; }

    /// <summary>Clockwise rotation in degrees.</summary>
    public double RotationDegrees { get; }

    /// <summary>Rotation and flip center X coordinate.</summary>
    public double RotationCenterX { get; }

    /// <summary>Rotation and flip center Y coordinate.</summary>
    public double RotationCenterY { get; }

    /// <summary>Whether to mirror horizontally around the rotation center.</summary>
    public bool FlipHorizontal { get; }

    /// <summary>Whether to mirror vertically around the rotation center.</summary>
    public bool FlipVertical { get; }

    /// <summary>Destination left coordinate.</summary>
    public double X => Placement.X;

    /// <summary>Destination top coordinate.</summary>
    public double Y => Placement.Y;

    /// <summary>Destination width.</summary>
    public double Width => Placement.Width;

    /// <summary>Destination height.</summary>
    public double Height => Placement.Height;

    /// <summary>Normalized visible source left coordinate.</summary>
    public double SourceLeft => SourceCrop.Left;

    /// <summary>Normalized visible source top coordinate.</summary>
    public double SourceTop => SourceCrop.Top;

    /// <summary>Normalized visible source width.</summary>
    public double SourceWidth => SourceCrop.VisibleWidth;

    /// <summary>Normalized visible source height.</summary>
    public double SourceHeight => SourceCrop.VisibleHeight;

    /// <summary>Whether the source image is cropped.</summary>
    public bool HasCrop => SourceCrop.HasCrop;

    /// <summary>Whether the projection applies rotation.</summary>
    public bool HasRotation => Math.Abs(RotationDegrees) >= 0.000001D;

    /// <summary>Whether the projection applies rotation or flipping.</summary>
    public bool HasTransform => HasRotation || FlipHorizontal || FlipVertical;

    /// <summary>
    /// Creates a destination-space frame transform for this projection's rotation and mirror settings.
    /// </summary>
    /// <returns>A shared frame transform that can be converted to format-specific transform syntax or matrices.</returns>
    public OfficeImageFrameTransform CreateFrameTransform() =>
        new OfficeImageFrameTransform(
            RotationDegrees,
            RotationCenterX,
            RotationCenterY,
            FlipHorizontal,
            FlipVertical);

    /// <summary>
    /// Returns a projection scaled by the supplied output factor.
    /// </summary>
    /// <param name="scale">Positive scale factor.</param>
    /// <returns>A scaled projection with source crop and transform flags preserved.</returns>
    public OfficeImageProjection Scale(double scale) {
        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            throw new ArgumentOutOfRangeException(nameof(scale), "Image projection scale must be positive and finite.");
        }

        return new OfficeImageProjection(
            new OfficeImagePlacement(X * scale, Y * scale, Width * scale, Height * scale),
            SourceCrop,
            RotationDegrees,
            RotationCenterX * scale,
            RotationCenterY * scale,
            FlipHorizontal,
            FlipVertical);
    }

    /// <summary>
    /// Returns a projection translated by the supplied destination offset.
    /// </summary>
    /// <param name="offsetX">Horizontal destination offset.</param>
    /// <param name="offsetY">Vertical destination offset.</param>
    /// <returns>A translated projection with source crop and transform flags preserved.</returns>
    public OfficeImageProjection Translate(double offsetX, double offsetY) {
        EnsureFinite(offsetX, nameof(offsetX));
        EnsureFinite(offsetY, nameof(offsetY));

        return new OfficeImageProjection(
            new OfficeImagePlacement(X + offsetX, Y + offsetY, Width, Height),
            SourceCrop,
            RotationDegrees,
            RotationCenterX + offsetX,
            RotationCenterY + offsetY,
            FlipHorizontal,
            FlipVertical);
    }

    /// <summary>
    /// Creates an affine transform that maps a unit image rectangle into this projection's destination placement.
    /// </summary>
    /// <remarks>
    /// The returned matrix maps source coordinates from <c>0..1</c> into the destination coordinate system,
    /// applying rotation and flips around <see cref="RotationCenterX" /> / <see cref="RotationCenterY" />.
    /// </remarks>
    public OfficeTransform CreateUnitSquareTransform() {
        double radians = RotationDegrees * Math.PI / 180D;
        double cos = NormalizeZero(Math.Cos(radians));
        double sin = NormalizeZero(Math.Sin(radians));

        double a = NormalizeZero(Width * cos);
        double b = NormalizeZero(Width * sin);
        double c = NormalizeZero(-Height * sin);
        double d = NormalizeZero(Height * cos);
        double e = NormalizeZero(RotationCenterX + (cos * (X - RotationCenterX)) - (sin * (Y - RotationCenterY)));
        double f = NormalizeZero(RotationCenterY + (sin * (X - RotationCenterX)) + (cos * (Y - RotationCenterY)));

        if (FlipHorizontal) {
            e = NormalizeZero(e + a);
            f = NormalizeZero(f + b);
            a = NormalizeZero(-a);
            b = NormalizeZero(-b);
        }

        if (FlipVertical) {
            e = NormalizeZero(e + c);
            f = NormalizeZero(f + d);
            c = NormalizeZero(-c);
            d = NormalizeZero(-d);
        }

        return new OfficeTransform(a, b, c, d, e, f);
    }

    /// <summary>
    /// Calculates axis-aligned destination bounds after applying this projection's rotation and mirror settings.
    /// </summary>
    /// <returns>Projected destination bounds.</returns>
    public (double Left, double Top, double Right, double Bottom) GetDestinationBounds() =>
        CreateUnitSquareTransform().TransformRectangleBounds(0D, 0D, 1D, 1D);

    /// <summary>
    /// Converts the projection into a tuple useful for assertions and diagnostics.
    /// </summary>
    public (double X, double Y, double Width, double Height, double SourceLeft, double SourceTop, double SourceWidth, double SourceHeight, double RotationDegrees, double RotationCenterX, double RotationCenterY, bool FlipHorizontal, bool FlipVertical) ToTuple() =>
        (X, Y, Width, Height, SourceLeft, SourceTop, SourceWidth, SourceHeight, RotationDegrees, RotationCenterX, RotationCenterY, FlipHorizontal, FlipVertical);

    private static double NormalizeZero(double value) => Math.Abs(value) < 0.000000000001D ? 0D : value;

    private static void EnsureFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Image projection transform values must be finite.");
        }
    }
}
