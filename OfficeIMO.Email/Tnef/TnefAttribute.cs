namespace OfficeIMO.Email;

/// <summary>Application level of a TNEF attribute.</summary>
public enum TnefAttributeLevel : byte {
    /// <summary>Message-level attribute.</summary>
    Message = 1,
    /// <summary>Attachment-level attribute.</summary>
    Attachment = 2
}

/// <summary>Raw, ordered TNEF attribute retained for diagnostics and regeneration.</summary>
public sealed class TnefAttribute {
    private readonly byte[] _data;

    /// <summary>Creates an attribute.</summary>
    public TnefAttribute(TnefAttributeLevel level, uint tag, byte[] data, bool checksumIsValid = true) {
        Level = level;
        Tag = tag;
        if (data == null) throw new ArgumentNullException(nameof(data));
        _data = (byte[])data.Clone();
        ChecksumIsValid = checksumIsValid;
    }

    /// <summary>Message or attachment level.</summary>
    public TnefAttributeLevel Level { get; }
    /// <summary>Combined TNEF attribute identifier and type.</summary>
    public uint Tag { get; }
    /// <summary>Attribute payload.</summary>
    public byte[] Data => (byte[])_data.Clone();
    /// <summary>Whether the stored checksum matched the payload.</summary>
    public bool ChecksumIsValid { get; }
}
