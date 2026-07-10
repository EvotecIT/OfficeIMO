namespace OfficeIMO.Email;

/// <summary>Typed MAPI property retained from or intended for an MSG/TNEF artifact.</summary>
public sealed class MapiProperty {
    /// <summary>Creates a property.</summary>
    public MapiProperty(ushort propertyId, MapiPropertyType propertyType, object? value = null,
        uint flags = 0x00000006, MapiNamedProperty? name = null) {
        PropertyId = propertyId;
        PropertyType = propertyType;
        Value = value;
        Flags = flags;
        Name = name;
    }

    /// <summary>Property ID encoded in the high 16 bits of a property tag.</summary>
    public ushort PropertyId { get; set; }

    /// <summary>Property value type.</summary>
    public MapiPropertyType PropertyType { get; set; }

    /// <summary>Combined property tag.</summary>
    public uint PropertyTag => ((uint)PropertyId << 16) | (ushort)PropertyType;

    /// <summary>Property stream flags.</summary>
    public uint Flags { get; set; }

    /// <summary>Decoded scalar or array value.</summary>
    public object? Value { get; set; }

    /// <summary>Original serialized value bytes when available.</summary>
    public byte[]? RawData { get; set; }

    /// <summary>Canonical named-property identity for mapped IDs.</summary>
    public MapiNamedProperty? Name { get; set; }
}
