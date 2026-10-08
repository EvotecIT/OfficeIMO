// Adapted from the Evotec ImagePlayground HEIF metadata implementation.
// Copyright (c) 2022 Evotec. MIT license; see Licenses/ImagePlayground-LICENSE.txt.
// OfficeIMO adaptation provides bounded byte/stream APIs, cancellation and output preflight.
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes high-level HEIF container metadata that can be read without decoding image pixels.
/// </summary>
public sealed class OfficeHeifImageInfo {
    internal OfficeHeifImageInfo(string majorBrand, uint minorVersion, IReadOnlyList<string> compatibleBrands, uint? primaryItemId, IReadOnlyList<OfficeHeifItemInfo> items, IReadOnlyList<OfficeHeifItemReference> references, uint? exifItemId, uint? xmpItemId) {
        MajorBrand = majorBrand;
        MinorVersion = minorVersion;
        CompatibleBrands = compatibleBrands;
        PrimaryItemId = primaryItemId;
        Items = items;
        References = references;
        PrimaryItem = primaryItemId.HasValue
            ? items.FirstOrDefault(item => item.ItemId == primaryItemId.Value)
            : null;
        ExifItem = exifItemId.HasValue ? items.FirstOrDefault(item => item.ItemId == exifItemId.Value) : null;
        XmpItem = xmpItemId.HasValue ? items.FirstOrDefault(item => item.ItemId == xmpItemId.Value) : null;
    }

    /// <summary>Major HEIF brand declared by the file type box.</summary>
    public string MajorBrand { get; }

    /// <summary>Minor file type version declared by the file type box.</summary>
    public uint MinorVersion { get; }

    /// <summary>Compatible brands declared by the file type box.</summary>
    public IReadOnlyList<string> CompatibleBrands { get; }

    /// <summary>Primary item identifier when the file declares one.</summary>
    public uint? PrimaryItemId { get; }

    /// <summary>Primary item metadata when it could be resolved from <see cref="Items"/>.</summary>
    public OfficeHeifItemInfo? PrimaryItem { get; }

    /// <summary>Primary-associated or unique EXIF item metadata; null when selection is ambiguous or unrelated to the primary item.</summary>
    public OfficeHeifItemInfo? ExifItem { get; }

    /// <summary>Primary-associated or unique XMP item metadata; null when selection is ambiguous or unrelated to the primary item.</summary>
    public OfficeHeifItemInfo? XmpItem { get; }

    /// <summary>Known item metadata declared by the HEIF item information box.</summary>
    public IReadOnlyList<OfficeHeifItemInfo> Items { get; }

    /// <summary>Item references declared by the HEIF item reference box.</summary>
    public IReadOnlyList<OfficeHeifItemReference> References { get; }

    /// <summary>Image width from the primary item when declared by HEIF item properties.</summary>
    public uint? Width => PrimaryItem?.Width;

    /// <summary>Image height from the primary item when declared by HEIF item properties.</summary>
    public uint? Height => PrimaryItem?.Height;

    /// <summary>Rotation angle from the primary item's HEIF <c>irot</c> property when present.</summary>
    public int? RotationDegrees => PrimaryItem?.RotationDegrees;

    /// <summary>Whether the primary item has an associated HEIF <c>imir</c> mirror property.</summary>
    public bool IsMirrored => PrimaryItem?.IsMirrored ?? false;

    /// <summary>Horizontal spacing from the primary item's HEIF <c>pasp</c> pixel aspect ratio property when present.</summary>
    public uint? PixelAspectRatioHorizontalSpacing => PrimaryItem?.PixelAspectRatioHorizontalSpacing;

    /// <summary>Vertical spacing from the primary item's HEIF <c>pasp</c> pixel aspect ratio property when present.</summary>
    public uint? PixelAspectRatioVerticalSpacing => PrimaryItem?.PixelAspectRatioVerticalSpacing;

    /// <summary>Bits per channel from the primary item's HEIF <c>pixi</c> pixel information property when present.</summary>
    public IReadOnlyList<byte> PixelBitDepths => PrimaryItem?.PixelBitDepths ?? Array.Empty<byte>();

    /// <summary>Color information type from the primary item's HEIF <c>colr</c> property when present.</summary>
    public string? ColorType => PrimaryItem?.ColorType;

    /// <summary>Color primaries value from the primary item's HEIF <c>colr</c> <c>nclx</c> property when present.</summary>
    public ushort? ColorPrimaries => PrimaryItem?.ColorPrimaries;

    /// <summary>Transfer characteristics value from the primary item's HEIF <c>colr</c> <c>nclx</c> property when present.</summary>
    public ushort? TransferCharacteristics => PrimaryItem?.TransferCharacteristics;

    /// <summary>Matrix coefficients value from the primary item's HEIF <c>colr</c> <c>nclx</c> property when present.</summary>
    public ushort? MatrixCoefficients => PrimaryItem?.MatrixCoefficients;

    /// <summary>Full-range flag from the primary item's HEIF <c>colr</c> <c>nclx</c> property when present.</summary>
    public bool? FullRangeFlag => PrimaryItem?.FullRangeFlag;

    /// <summary>Codec configuration property type from the primary item when present.</summary>
    public string? CodecConfigurationType => PrimaryItem?.CodecConfigurationType;

    /// <summary>Codec configuration payload bytes from the primary item when present.</summary>
    public IReadOnlyList<byte> CodecConfigurationBytes => PrimaryItem?.CodecConfigurationBytes ?? Array.Empty<byte>();

    /// <summary>Whether the container declares an EXIF metadata item.</summary>
    public bool HasExif => Items.Any(item => item.HasExif);

    /// <summary>Whether the container declares an XMP metadata item.</summary>
    public bool HasXmp => Items.Any(item => item.HasXmp);
}

/// <summary>
/// Describes a HEIF item declared in the item information box.
/// </summary>
public sealed class OfficeHeifItemInfo {
    internal OfficeHeifItemInfo(uint itemId, string itemType, string itemName, ushort itemProtectionIndex, bool isHidden, string? mimeType, string? contentEncoding, OfficeHeifItemLocationInfo? location, IReadOnlyList<OfficeHeifItemPropertyAssociationInfo> propertyAssociations, bool isPrimary, bool hasExif, bool hasXmp, uint? width, uint? height, int? rotationDegrees, bool isMirrored, uint? pixelAspectRatioHorizontalSpacing, uint? pixelAspectRatioVerticalSpacing, IReadOnlyList<byte> pixelBitDepths, string? colorType, ushort? colorPrimaries, ushort? transferCharacteristics, ushort? matrixCoefficients, bool? fullRangeFlag, string? codecConfigurationType, IReadOnlyList<byte> codecConfigurationBytes, string? auxiliaryType, IReadOnlyList<byte> auxiliarySubtypes) {
        ItemId = itemId;
        ItemType = itemType;
        ItemName = itemName;
        ItemProtectionIndex = itemProtectionIndex;
        IsHidden = isHidden;
        MimeType = mimeType;
        ContentEncoding = contentEncoding;
        Location = location;
        PropertyAssociations = propertyAssociations;
        IsPrimary = isPrimary;
        HasExif = hasExif;
        HasXmp = hasXmp;
        Width = width;
        Height = height;
        RotationDegrees = rotationDegrees;
        IsMirrored = isMirrored;
        PixelAspectRatioHorizontalSpacing = pixelAspectRatioHorizontalSpacing;
        PixelAspectRatioVerticalSpacing = pixelAspectRatioVerticalSpacing;
        PixelBitDepths = pixelBitDepths;
        ColorType = colorType;
        ColorPrimaries = colorPrimaries;
        TransferCharacteristics = transferCharacteristics;
        MatrixCoefficients = matrixCoefficients;
        FullRangeFlag = fullRangeFlag;
        CodecConfigurationType = codecConfigurationType;
        CodecConfigurationBytes = codecConfigurationBytes;
        AuxiliaryType = auxiliaryType;
        AuxiliarySubtypes = auxiliarySubtypes;
    }

    /// <summary>Item identifier used by HEIF references and item-location records.</summary>
    public uint ItemId { get; }

    /// <summary>Four-character item type such as <c>hvc1</c>, <c>av01</c>, or <c>Exif</c>.</summary>
    public string ItemType { get; }

    /// <summary>Optional item name from the item information entry.</summary>
    public string ItemName { get; }

    /// <summary>Item protection index from the HEIF item information entry.</summary>
    public ushort ItemProtectionIndex { get; }

    /// <summary>Whether the item information entry marks this item as hidden.</summary>
    public bool IsHidden { get; }

    /// <summary>MIME content type for <c>mime</c> items when present.</summary>
    public string? MimeType { get; }

    /// <summary>MIME content encoding for <c>mime</c> items when present.</summary>
    public string? ContentEncoding { get; }

    /// <summary>Item location metadata from the HEIF <c>iloc</c> box when present.</summary>
    public OfficeHeifItemLocationInfo? Location { get; }

    /// <summary>HEIF item property associations from the <c>ipma</c> box.</summary>
    public IReadOnlyList<OfficeHeifItemPropertyAssociationInfo> PropertyAssociations { get; }

    /// <summary>Whether this item is declared as the primary item.</summary>
    public bool IsPrimary { get; }

    /// <summary>Whether this item is the EXIF metadata item.</summary>
    public bool HasExif { get; }

    /// <summary>Whether this item is an XMP metadata item.</summary>
    public bool HasXmp { get; }

    /// <summary>Image width from the HEIF <c>ispe</c> property when present.</summary>
    public uint? Width { get; }

    /// <summary>Image height from the HEIF <c>ispe</c> property when present.</summary>
    public uint? Height { get; }

    /// <summary>Rotation angle from the HEIF <c>irot</c> property when present.</summary>
    public int? RotationDegrees { get; }

    /// <summary>Whether the HEIF <c>imir</c> property is associated with this item.</summary>
    public bool IsMirrored { get; }

    /// <summary>Horizontal spacing from the HEIF <c>pasp</c> pixel aspect ratio property when present.</summary>
    public uint? PixelAspectRatioHorizontalSpacing { get; }

    /// <summary>Vertical spacing from the HEIF <c>pasp</c> pixel aspect ratio property when present.</summary>
    public uint? PixelAspectRatioVerticalSpacing { get; }

    /// <summary>Bits per channel from the HEIF <c>pixi</c> pixel information property when present.</summary>
    public IReadOnlyList<byte> PixelBitDepths { get; }

    /// <summary>Color information type from the HEIF <c>colr</c> property when present, such as <c>nclx</c>, <c>rICC</c>, or <c>prof</c>.</summary>
    public string? ColorType { get; }

    /// <summary>Color primaries value from an HEIF <c>colr</c> <c>nclx</c> property when present.</summary>
    public ushort? ColorPrimaries { get; }

    /// <summary>Transfer characteristics value from an HEIF <c>colr</c> <c>nclx</c> property when present.</summary>
    public ushort? TransferCharacteristics { get; }

    /// <summary>Matrix coefficients value from an HEIF <c>colr</c> <c>nclx</c> property when present.</summary>
    public ushort? MatrixCoefficients { get; }

    /// <summary>Full-range flag from an HEIF <c>colr</c> <c>nclx</c> property when present.</summary>
    public bool? FullRangeFlag { get; }

    /// <summary>Codec configuration property type when present, such as <c>hvcC</c>, <c>av1C</c>, or <c>avcC</c>.</summary>
    public string? CodecConfigurationType { get; }

    /// <summary>Raw codec configuration payload bytes when present.</summary>
    public IReadOnlyList<byte> CodecConfigurationBytes { get; }

    /// <summary>Auxiliary image type URI from the HEIF <c>auxC</c> property when present.</summary>
    public string? AuxiliaryType { get; }

    /// <summary>Auxiliary subtype bytes from the HEIF <c>auxC</c> property when present.</summary>
    public IReadOnlyList<byte> AuxiliarySubtypes { get; }
}

/// <summary>
/// Describes one HEIF item property association from an <c>ipma</c> box.
/// </summary>
public sealed class OfficeHeifItemPropertyAssociationInfo {
    internal OfficeHeifItemPropertyAssociationInfo(int propertyIndex, string propertyType, bool isEssential) {
        PropertyIndex = propertyIndex;
        PropertyType = propertyType;
        IsEssential = isEssential;
    }

    /// <summary>One-based property index in the item property container.</summary>
    public int PropertyIndex { get; }

    /// <summary>Four-character property type, such as <c>ispe</c>, <c>irot</c>, <c>colr</c>, or <c>auxC</c>.</summary>
    public string PropertyType { get; }

    /// <summary>Whether the association marks this property as essential for the item.</summary>
    public bool IsEssential { get; }
}

/// <summary>
/// Describes where a HEIF item payload is stored.
/// </summary>
public sealed class OfficeHeifItemLocationInfo {
    internal OfficeHeifItemLocationInfo(ushort constructionMethod, ushort dataReferenceIndex, ulong baseOffset, IReadOnlyList<OfficeHeifItemExtentInfo> extents) {
        ConstructionMethod = constructionMethod;
        DataReferenceIndex = dataReferenceIndex;
        BaseOffset = baseOffset;
        Extents = extents;
    }

    /// <summary>HEIF item construction method from the item location box.</summary>
    public ushort ConstructionMethod { get; }

    /// <summary>HEIF data reference index from the item location box.</summary>
    public ushort DataReferenceIndex { get; }

    /// <summary>Base offset used to resolve item extents.</summary>
    public ulong BaseOffset { get; }

    /// <summary>Resolved extents declared for the item.</summary>
    public IReadOnlyList<OfficeHeifItemExtentInfo> Extents { get; }

    /// <summary>Whether the item is stored in the same file rather than an external data reference or derived construction.</summary>
    public bool IsFileBacked => ConstructionMethod == 0 && DataReferenceIndex == 0;

    /// <summary>Whether the item is stored inside the HEIF metadata <c>idat</c> box.</summary>
    public bool IsItemDataBoxBacked => ConstructionMethod == 1 && DataReferenceIndex == 0;

    /// <summary>Whether the current writer can safely update this item by replacing one file-backed extent.</summary>
    public bool CanWriteSingleFileExtent => IsFileBacked && Extents.Count == 1;
}

/// <summary>
/// Describes one HEIF item extent.
/// </summary>
public sealed class OfficeHeifItemExtentInfo {
    internal OfficeHeifItemExtentInfo(int offset, int length) {
        Offset = offset;
        Length = length;
    }

    /// <summary>Resolved extent offset.</summary>
    public int Offset { get; }

    /// <summary>Extent length in bytes.</summary>
    public int Length { get; }
}

/// <summary>
/// Describes a HEIF item reference.
/// </summary>
public sealed class OfficeHeifItemReference {
    internal OfficeHeifItemReference(string referenceType, uint fromItemId, IReadOnlyList<uint> toItemIds) {
        ReferenceType = referenceType;
        FromItemId = fromItemId;
        ToItemIds = toItemIds;
    }

    /// <summary>Four-character reference type such as <c>cdsc</c> or <c>thmb</c>.</summary>
    public string ReferenceType { get; }

    /// <summary>Source item identifier for this reference.</summary>
    public uint FromItemId { get; }

    /// <summary>Target item identifiers for this reference.</summary>
    public IReadOnlyList<uint> ToItemIds { get; }
}
