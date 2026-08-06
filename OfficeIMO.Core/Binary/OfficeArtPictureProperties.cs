using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

/// <summary>
/// Decodes OfficeArt picture-frame properties that can be shared by binary Office consumers.
/// </summary>
public sealed class OfficeArtPictureProperties {
    private OfficeArtPictureProperties(IReadOnlyList<OfficeArtProperty> properties) {
        Properties = properties?.ToArray() ?? Array.Empty<OfficeArtProperty>();
        CropFromTopRaw = GetSignedValue(0x0100);
        CropFromBottomRaw = GetSignedValue(0x0101);
        CropFromLeftRaw = GetSignedValue(0x0102);
        CropFromRightRaw = GetSignedValue(0x0103);
        TransparentColor = GetColor(0x0107);
        ContrastRaw = GetSignedValue(0x0108);
        BrightnessRaw = GetSignedValue(0x0109);
        RecolorColor = GetColor(0x011A);
        Grayscale = GetBoolean(0x013F, 1U << 18, 1U << 2);
        BiLevel = GetBoolean(0x013F, 1U << 17, 1U << 1);
    }

    /// <summary>Decodes picture-frame properties from an OfficeArt property table.</summary>
    public static OfficeArtPictureProperties Decode(IReadOnlyList<OfficeArtProperty>? properties) =>
        new OfficeArtPictureProperties(properties ?? Array.Empty<OfficeArtProperty>());

    /// <summary>Gets the source property entries.</summary>
    public IReadOnlyList<OfficeArtProperty> Properties { get; }

    /// <summary>Gets the raw signed 16.16 top crop value.</summary>
    public int? CropFromTopRaw { get; }

    /// <summary>Gets the raw signed 16.16 bottom crop value.</summary>
    public int? CropFromBottomRaw { get; }

    /// <summary>Gets the raw signed 16.16 left crop value.</summary>
    public int? CropFromLeftRaw { get; }

    /// <summary>Gets the raw signed 16.16 right crop value.</summary>
    public int? CropFromRightRaw { get; }

    /// <summary>Gets the color treated as transparent, when explicitly present.</summary>
    public OfficeArtColorReference? TransparentColor { get; }

    /// <summary>Gets the raw signed picture-contrast value, where 0x00010000 means unchanged.</summary>
    public int? ContrastRaw { get; }

    /// <summary>Gets the raw signed picture-brightness value in the range -0x8000 through 0x8000.</summary>
    public int? BrightnessRaw { get; }

    /// <summary>
    /// Gets brightness adjustment from -1 through 1 by scaling the signed OfficeArt endpoint
    /// range of -0x8000 through 0x8000.
    /// </summary>
    public double? BrightnessAdjustment => BrightnessRaw.HasValue
        ? Clamp(BrightnessRaw.Value / 32768D)
        : null;

    /// <summary>
    /// Gets contrast adjustment from -1 through 1. Values through 0x10000 scale toward minimum
    /// contrast; larger multipliers scale asymptotically toward maximum contrast.
    /// </summary>
    public double? ContrastAdjustment {
        get {
            if (!ContrastRaw.HasValue || ContrastRaw.Value < 0) return null;
            if (ContrastRaw.Value == int.MaxValue) return 1D;
            return ContrastRaw.Value <= 0x10000
                ? Clamp(ContrastRaw.Value / 65536D - 1D)
                : Clamp(1D - 65536D / ContrastRaw.Value);
        }
    }

    /// <summary>Gets the picture recolor color, when explicitly present.</summary>
    public OfficeArtColorReference? RecolorColor { get; }

    /// <summary>Gets explicit grayscale display state, or null when inherited.</summary>
    public bool? Grayscale { get; }

    /// <summary>Gets explicit two-color black-and-white display state, or null when inherited.</summary>
    public bool? BiLevel { get; }

    /// <summary>Gets the top crop as a fraction of image height.</summary>
    public double? CropFromTop => ToFraction(CropFromTopRaw);

    /// <summary>Gets the bottom crop as a fraction of image height.</summary>
    public double? CropFromBottom => ToFraction(CropFromBottomRaw);

    /// <summary>Gets the left crop as a fraction of image width.</summary>
    public double? CropFromLeft => ToFraction(CropFromLeftRaw);

    /// <summary>Gets the right crop as a fraction of image width.</summary>
    public double? CropFromRight => ToFraction(CropFromRightRaw);

    /// <summary>Gets whether at least one explicit crop property is present.</summary>
    public bool HasExplicitCrop => CropFromTopRaw.HasValue || CropFromBottomRaw.HasValue
        || CropFromLeftRaw.HasValue || CropFromRightRaw.HasValue;

    /// <summary>Gets whether any edge crops into or out from the source image.</summary>
    public bool HasCrop => CropFromTopRaw.GetValueOrDefault() != 0
        || CropFromBottomRaw.GetValueOrDefault() != 0
        || CropFromLeftRaw.GetValueOrDefault() != 0
        || CropFromRightRaw.GetValueOrDefault() != 0;

    /// <summary>Gets whether the picture has an explicitly enabled visual effect.</summary>
    public bool HasPictureEffect => TransparentColor.HasValue && !TransparentColor.Value.IsIgnored
        || ContrastRaw.HasValue && ContrastRaw.Value != 0x10000
        || BrightnessRaw.GetValueOrDefault() != 0
        || RecolorColor.HasValue && !RecolorColor.Value.IsIgnored
        || Grayscale == true || BiLevel == true;

    private int? GetSignedValue(ushort propertyId) {
        OfficeArtProperty? property = Properties.LastOrDefault(item =>
            item.PropertyId == propertyId && !item.IsComplex);
        return property == null ? null : unchecked((int)property.Value);
    }

    private OfficeArtColorReference? GetColor(ushort propertyId) {
        OfficeArtProperty? property = Properties.LastOrDefault(item =>
            item.PropertyId == propertyId && !item.IsComplex);
        return property == null ? null : new OfficeArtColorReference(property.Value);
    }

    private bool? GetBoolean(ushort propertyId, uint useMask, uint valueMask) {
        OfficeArtProperty? property = Properties.LastOrDefault(item =>
            item.PropertyId == propertyId && !item.IsComplex);
        if (property == null || (property.Value & useMask) == 0) return null;
        return (property.Value & valueMask) != 0;
    }

    private static double? ToFraction(int? value) => value.HasValue
        ? value.Value / 65536D
        : null;

    private static double Clamp(double value) => Math.Max(-1D, Math.Min(1D, value));
}
