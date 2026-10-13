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
    private const uint FlipHorizontalOverride = 1U << 24;
    private const uint FlipVerticalOverride = 1U << 25;
    private const uint UseFlipHorizontalOverride = 1U << 8;
    private const uint UseFlipVerticalOverride = 1U << 9;

    private OfficeArtShapeTransform(double? rotationDegrees, bool flipHorizontal, bool flipVertical) {
        RotationDegrees = rotationDegrees;
        FlipHorizontal = flipHorizontal;
        FlipVertical = flipVertical;
    }

    /// <summary>Decodes transform state from FSP flags and OfficeArt properties.</summary>
    public static OfficeArtShapeTransform Decode(uint fspFlags,
        IReadOnlyList<OfficeArtProperty>? properties = null) => Decode(fspFlags, properties, false);

    // RelXfrm belongs to a parent coordinate space; callers opt in only after
    // identifying group membership, independently of the anchor encoding.
    // Explicit relative zero overrides absolute rotation.
    internal static OfficeArtShapeTransform Decode(uint fspFlags,
        IReadOnlyList<OfficeArtProperty>? properties, bool relativeToParent) {
        IReadOnlyList<OfficeArtProperty> source = properties ?? Array.Empty<OfficeArtProperty>();
        OfficeArtProperty? rotation = OfficeArtProperty.GetScalarProperty(source,
            relativeToParent && source.Any(property => property.PropertyId == 0x03C4) ? 0x03C4 : 0x0004);
        double? rotationDegrees = rotation == null
            ? null
            : unchecked((int)rotation.Value) / 65536D;

        bool flipHorizontal = OfficeArtProperty.GetBoolean(source, 0x033F, UseFlipHorizontalOverride, FlipHorizontalOverride)
            ?? (fspFlags & FlipHorizontalFlag) != 0;
        bool flipVertical = OfficeArtProperty.GetBoolean(source, 0x033F, UseFlipVerticalOverride, FlipVerticalOverride)
            ?? (fspFlags & FlipVerticalFlag) != 0;

        if (relativeToParent) {
            flipHorizontal = OfficeArtProperty.GetBoolean(source, 0x03FF, 1U << 15, 1U << 31) ?? flipHorizontal;
            flipVertical = OfficeArtProperty.GetBoolean(source, 0x03FF, 1U << 14, 1U << 30) ?? flipVertical;
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
