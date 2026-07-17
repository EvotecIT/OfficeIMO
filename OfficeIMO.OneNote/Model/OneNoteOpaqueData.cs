namespace OfficeIMO.OneNote;

/// <summary>
/// An unsupported property retained byte-for-byte for diagnostics and round-trip writing.
/// </summary>
public sealed class OneNoteOpaqueProperty {
    private byte[] _rawData = Array.Empty<byte>();

    /// <summary>MS-ONESTORE property identifier.</summary>
    public uint PropertyId { get; set; }

    /// <summary>Decoded representation category when known.</summary>
    public OneNotePropertyValueType ValueType { get; set; }

    /// <summary>Original ordinal within the property set.</summary>
    public int Ordinal { get; set; }

    /// <summary>Inline Boolean value, when this is a Boolean property.</summary>
    public bool? BooleanValue { get; set; }

    /// <summary>Unsigned scalar value, when this is a fixed-width numeric property.</summary>
    public ulong? ScalarValue { get; set; }

    /// <summary>Resolved object, object-space, or context references.</summary>
    public IList<OneNoteExtendedGuid> ReferencedIds { get; } = new List<OneNoteExtendedGuid>();

    /// <summary>Replaces the retained encoded data with a defensive copy.</summary>
    public void SetRawData(byte[] rawData) {
        if (rawData == null) throw new ArgumentNullException(nameof(rawData));
        _rawData = new byte[rawData.Length];
        Buffer.BlockCopy(rawData, 0, _rawData, 0, rawData.Length);
    }

    /// <summary>Returns a defensive copy of the retained encoded data.</summary>
    public byte[] GetRawData() {
        var copy = new byte[_rawData.Length];
        Buffer.BlockCopy(_rawData, 0, copy, 0, copy.Length);
        return copy;
    }
}

/// <summary>
/// An unsupported object retained byte-for-byte for diagnostics and round-trip writing.
/// </summary>
public sealed class OneNoteOpaqueObject {
    private byte[] _rawData = Array.Empty<byte>();

    /// <summary>Object identifier when available. Serialization assigns and retains one for a new opaque object.</summary>
    public OneNoteExtendedGuid? Id { get; set; }

    /// <summary>JCID value that identifies the object type.</summary>
    public uint Jcid { get; set; }

    /// <summary>Original ordinal within its object group.</summary>
    public int Ordinal { get; set; }

    /// <summary>Unsupported properties decoded from the object.</summary>
    public IList<OneNoteOpaqueProperty> Properties { get; } = new List<OneNoteOpaqueProperty>();

    /// <summary>Replaces the retained encoded object data with a defensive copy.</summary>
    public void SetRawData(byte[] rawData) {
        if (rawData == null) throw new ArgumentNullException(nameof(rawData));
        _rawData = new byte[rawData.Length];
        Buffer.BlockCopy(rawData, 0, _rawData, 0, rawData.Length);
    }

    /// <summary>Returns a defensive copy of the retained encoded object data.</summary>
    public byte[] GetRawData() {
        var copy = new byte[_rawData.Length];
        Buffer.BlockCopy(_rawData, 0, copy, 0, copy.Length);
        return copy;
    }
}
