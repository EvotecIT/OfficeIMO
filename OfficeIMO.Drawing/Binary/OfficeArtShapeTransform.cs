using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

/// <summary>
/// Decodes rotation and mirroring from an OfficeArt FSP record and its primary property table.
/// </summary>
public sealed class OfficeArtShapeTransform {
    private const uint FlipHorizontalFlag = 1U << 6;
    private const uint FlipVerticalFlag = 1U << 7;
    private const uint UseFlipHorizontalOverride = 1U << 8;
    private const uint UseFlipVerticalOverride = 1U << 9;
    private const uint FlipHorizontalOverride = 1U << 24;
    private const uint FlipVerticalOverride = 1U << 25;

    private OfficeArtShapeTransform(double? rotationDegrees, bool flipHorizontal, bool flipVertical) {
        RotationDegrees = rotationDegrees;
        FlipHorizontal = flipHorizontal;
        FlipVertical = flipVertical;
    }

    /// <summary>Decodes transform state from FSP flags and OfficeArt properties.</summary>
    public static OfficeArtShapeTransform Decode(uint fspFlags,
        IReadOnlyList<OfficeArtProperty>? properties = null) {
        IReadOnlyList<OfficeArtProperty> source = properties ?? Array.Empty<OfficeArtProperty>();
        OfficeArtProperty? rotation = source.LastOrDefault(property =>
            property.PropertyId == 0x0004 && !property.IsComplex);
        double? rotationDegrees = rotation == null
            ? null
            : unchecked((int)rotation.Value) / 65536D;

        bool flipHorizontal = (fspFlags & FlipHorizontalFlag) != 0;
        bool flipVertical = (fspFlags & FlipVerticalFlag) != 0;
        OfficeArtProperty? shapeBooleans = source.LastOrDefault(property =>
            property.PropertyId == 0x033F && !property.IsComplex);
        if (shapeBooleans != null) {
            uint value = shapeBooleans.Value;
            if ((value & UseFlipHorizontalOverride) != 0) {
                flipHorizontal = (value & FlipHorizontalOverride) != 0;
            }
            if ((value & UseFlipVerticalOverride) != 0) {
                flipVertical = (value & FlipVerticalOverride) != 0;
            }
        }

        return new OfficeArtShapeTransform(rotationDegrees, flipHorizontal, flipVertical);
    }

    /// <summary>Gets clockwise rotation in degrees, or null when no rotation property is present.</summary>
    public double? RotationDegrees { get; }

    /// <summary>Gets whether the shape is mirrored horizontally.</summary>
    public bool FlipHorizontal { get; }

    /// <summary>Gets whether the shape is mirrored vertically.</summary>
    public bool FlipVertical { get; }

    /// <summary>Gets whether this transform carries rotation or mirroring.</summary>
    public bool HasTransform => RotationDegrees.HasValue || FlipHorizontal || FlipVertical;
}
