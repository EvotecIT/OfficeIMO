using System;

namespace OfficeIMO.Drawing.Binary;

/// <summary>
/// Describes one property entry from an OfficeArt FOPT property table.
/// </summary>
public sealed class OfficeArtProperty {
    /// <summary>Creates an OfficeArt property entry.</summary>
    public OfficeArtProperty(int index, ushort rawOperationId, uint value,
        int? availableComplexDataLength = null, string? complexText = null,
        byte[]? complexData = null) {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        if (availableComplexDataLength < 0) {
            throw new ArgumentOutOfRangeException(nameof(availableComplexDataLength));
        }

        Index = index;
        RawOperationId = rawOperationId;
        PropertyId = checked((ushort)(rawOperationId & 0x3fff));
        PropertyName = GetPropertyName(PropertyId);
        PropertyGroupName = GetPropertyGroupName(PropertyId);
        IsBlipId = (rawOperationId & 0x4000) != 0;
        IsComplex = (rawOperationId & 0x8000) != 0;
        Value = value;
        AvailableComplexDataLength = availableComplexDataLength;
        ComplexText = string.IsNullOrWhiteSpace(complexText) ? null : complexText;
        _complexData = complexData == null ? null : (byte[])complexData.Clone();
    }

    private readonly byte[]? _complexData;

    /// <summary>Gets the zero-based index inside the property table.</summary>
    public int Index { get; }

    /// <summary>Gets the raw OfficeArtFOPTEOPID bitfield.</summary>
    public ushort RawOperationId { get; }

    /// <summary>Gets the low 14-bit OfficeArt property identifier.</summary>
    public ushort PropertyId { get; }

    /// <summary>Gets a stable property identifier display key.</summary>
    public string PropertyIdKey => $"PropertyId:0x{PropertyId:X4}";

    /// <summary>Gets the decoded property name, or a stable identifier for an unknown property.</summary>
    public string PropertyName { get; }

    /// <summary>Gets the decoded OfficeArt property family.</summary>
    public string PropertyGroupName { get; }

    /// <summary>Gets whether the property value references BLIP data.</summary>
    public bool IsBlipId { get; }

    /// <summary>Gets whether the property has a following complex-data payload.</summary>
    public bool IsComplex { get; }

    /// <summary>Gets the raw 32-bit property value.</summary>
    public uint Value { get; }

    /// <summary>Gets the declared complex-data length.</summary>
    public uint? DeclaredComplexDataLength => IsComplex ? Value : null;

    /// <summary>Gets the number of complex-data bytes available in the containing record.</summary>
    public int? AvailableComplexDataLength { get; }

    /// <summary>Gets decoded text for a text-bearing complex property.</summary>
    public string? ComplexText { get; }

    /// <summary>Returns a defensive copy of the complex-data payload, when present.</summary>
    public byte[]? CopyComplexData() => _complexData == null ? null : (byte[])_complexData.Clone();

    private static string GetPropertyName(ushort propertyId) => propertyId switch {
        0x007F => "ProtectionBooleanProperties",
        0x00BF => "TextBooleanProperties",
        0x0104 => "pib",
        0x013F => "BlipBooleanProperties",
        0x0181 => "fillColor",
        0x0183 => "fillBackColor",
        0x01BF => "FillStyleBooleanProperties",
        0x01C0 => "lineColor",
        0x01CB => "lineWidth",
        0x01FF => "LineStyleBooleanProperties",
        0x023F => "ShadowStyleBooleanProperties",
        0x033F => "ShapeBooleanProperties",
        0x0380 => "wzName",
        0x03BF => "GroupShapeBooleanProperties",
        _ => $"PropertyId:0x{propertyId:X4}"
    };

    private static string GetPropertyGroupName(ushort propertyId) => propertyId switch {
        >= 0x0000 and <= 0x007F => "Protection",
        >= 0x0080 and <= 0x00BF => "Text",
        >= 0x0100 and <= 0x013F => "Blip",
        >= 0x0140 and <= 0x017F => "Geometry",
        >= 0x0180 and <= 0x01BF => "Fill",
        >= 0x01C0 and <= 0x01FF => "Line",
        >= 0x0200 and <= 0x023F => "Shadow",
        >= 0x0300 and <= 0x033F => "Shape",
        >= 0x0380 and <= 0x03BF => "GroupShape",
        _ => "Unknown"
    };
}
