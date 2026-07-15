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

    private int? GetSignedValue(ushort propertyId) {
        OfficeArtProperty? property = Properties.LastOrDefault(item =>
            item.PropertyId == propertyId && !item.IsComplex);
        return property == null ? null : unchecked((int)property.Value);
    }

    private static double? ToFraction(int? value) => value.HasValue
        ? value.Value / 65536D
        : null;
}
