using System;

namespace OfficeIMO.Drawing.Binary;

/// <summary>
/// Describes an OfficeArt File BLIP Store Entry and its bounded embedded or delayed image payload.
/// </summary>
public sealed class OfficeArtBlipStoreEntry {
    private readonly byte[] _imageBytes;

    internal OfficeArtBlipStoreEntry(ushort recordInstance, byte win32BlipType, byte macOsBlipType,
        string uidHex, ushort tag, uint sizeBytes, uint referenceCount, uint delayedStreamOffset,
        byte nameByteCount, string? name, OfficeArtBlipStorage storage, byte? blipRecordVersion,
        ushort? blipRecordInstance, ushort? blipRecordType, uint? blipPayloadLength,
        int? blipPayloadAvailableLength, string? blipPayloadSha256,
        byte[]? imageBytes, bool wasImageRejectedBySizeLimit) {
        RecordInstance = recordInstance;
        RecordInstanceBlipType = TryGetBlipType(recordInstance);
        RecordInstanceBlipTypeName = GetBlipTypeName(recordInstance);
        Win32BlipType = win32BlipType;
        Win32BlipTypeKind = TryGetBlipType(win32BlipType);
        Win32BlipTypeName = GetBlipTypeName(win32BlipType);
        MacOsBlipType = macOsBlipType;
        MacOsBlipTypeKind = TryGetBlipType(macOsBlipType);
        MacOsBlipTypeName = GetBlipTypeName(macOsBlipType);
        UidHex = uidHex;
        Tag = tag;
        SizeBytes = sizeBytes;
        ReferenceCount = referenceCount;
        DelayedStreamOffset = delayedStreamOffset;
        NameByteCount = nameByteCount;
        Name = name;
        Storage = storage;
        BlipRecordVersion = blipRecordVersion;
        BlipRecordInstance = blipRecordInstance;
        BlipRecordType = blipRecordType;
        BlipRecordTypeName = GetBlipRecordTypeName(blipRecordType);
        BlipPayloadLength = blipPayloadLength;
        BlipPayloadAvailableLength = blipPayloadAvailableLength;
        BlipPayloadSha256 = blipPayloadSha256;
        ContentType = GetContentType(blipRecordType, RecordInstanceBlipType,
            Win32BlipTypeKind, MacOsBlipTypeKind);
        _imageBytes = imageBytes == null ? Array.Empty<byte>() : (byte[])imageBytes.Clone();
        WasImageRejectedBySizeLimit = wasImageRejectedBySizeLimit;
    }

    /// <summary>Gets the BLIP type value stored in the FBSE OfficeArt record instance field.</summary>
    public ushort RecordInstance { get; }

    /// <summary>Gets the typed BLIP value from <see cref="RecordInstance"/>, when known.</summary>
    public OfficeArtBlipType? RecordInstanceBlipType { get; }

    /// <summary>Gets a stable display name for the FBSE record-instance BLIP type.</summary>
    public string RecordInstanceBlipTypeName { get; }

    /// <summary>Gets the Windows BLIP type byte.</summary>
    public byte Win32BlipType { get; }

    /// <summary>Gets the typed Windows BLIP value, when known.</summary>
    public OfficeArtBlipType? Win32BlipTypeKind { get; }

    /// <summary>Gets a stable display name for the Windows BLIP type.</summary>
    public string Win32BlipTypeName { get; }

    /// <summary>Gets the Macintosh BLIP type byte.</summary>
    public byte MacOsBlipType { get; }

    /// <summary>Gets the typed Macintosh BLIP value, when known.</summary>
    public OfficeArtBlipType? MacOsBlipTypeKind { get; }

    /// <summary>Gets a stable display name for the Macintosh BLIP type.</summary>
    public string MacOsBlipTypeName { get; }

    /// <summary>Gets the FBSE pixel-data UID as uppercase hexadecimal text.</summary>
    public string UidHex { get; }

    /// <summary>Gets the application-defined FBSE resource tag.</summary>
    public ushort Tag { get; }

    /// <summary>Gets the stored BLIP size in bytes.</summary>
    public uint SizeBytes { get; }

    /// <summary>Gets the number of references to the BLIP.</summary>
    public uint ReferenceCount { get; }

    /// <summary>Gets the associated delay-stream offset, or 0xFFFFFFFF when no delayed BLIP is declared.</summary>
    public uint DelayedStreamOffset { get; }

    /// <summary>Gets the declared UTF-16 name-data length in bytes.</summary>
    public byte NameByteCount { get; }

    /// <summary>Gets the optional decoded BLIP name.</summary>
    public string? Name { get; }

    /// <summary>Gets where the BLIP record was resolved.</summary>
    public OfficeArtBlipStorage Storage { get; }

    /// <summary>Gets the resolved BLIP OfficeArt record version.</summary>
    public byte? BlipRecordVersion { get; }

    /// <summary>Gets the resolved BLIP OfficeArt record instance.</summary>
    public ushort? BlipRecordInstance { get; }

    /// <summary>Gets the resolved BLIP OfficeArt record type.</summary>
    public ushort? BlipRecordType { get; }

    /// <summary>Gets a stable display name for the resolved BLIP record type.</summary>
    public string? BlipRecordTypeName { get; }

    /// <summary>Gets the declared BLIP record payload length.</summary>
    public uint? BlipPayloadLength { get; }

    /// <summary>Gets the number of BLIP payload bytes available inside the bounded source.</summary>
    public int? BlipPayloadAvailableLength { get; }

    /// <summary>Gets the SHA-256 hash of the available raw BLIP payload.</summary>
    public string? BlipPayloadSha256 { get; }

    /// <summary>Gets the image content type inferred from the BLIP record and FBSE types.</summary>
    public string? ContentType { get; }

    /// <summary>Gets whether the resolved BLIP payload is shorter than its declared length.</summary>
    public bool IsPayloadTruncated => BlipPayloadLength.HasValue && BlipPayloadAvailableLength.HasValue
        && BlipPayloadLength.Value > unchecked((uint)BlipPayloadAvailableLength.Value);

    /// <summary>Gets a defensive copy of image bytes suitable for an Open XML image part.</summary>
    public byte[] ImageBytes => (byte[])_imageBytes.Clone();

    /// <summary>Gets the number of decoded image bytes retained by this entry.</summary>
    public int ImageByteCount => _imageBytes.Length;

    /// <summary>Gets whether this entry can be projected as an Open XML image.</summary>
    public bool HasImportableImage => _imageBytes.Length > 0 && !string.IsNullOrWhiteSpace(ContentType);

    internal bool WasImageRejectedBySizeLimit { get; }

    internal static OfficeArtBlipType? TryGetBlipType(ushort value) => value switch {
        0x00 => OfficeArtBlipType.Error,
        0x01 => OfficeArtBlipType.Unknown,
        0x02 => OfficeArtBlipType.Emf,
        0x03 => OfficeArtBlipType.Wmf,
        0x04 => OfficeArtBlipType.Pict,
        0x05 => OfficeArtBlipType.Jpeg,
        0x06 => OfficeArtBlipType.Png,
        0x07 => OfficeArtBlipType.Dib,
        0x11 => OfficeArtBlipType.Tiff,
        0x12 => OfficeArtBlipType.CmykJpeg,
        _ => null
    };

    internal static string GetBlipTypeName(ushort value) =>
        TryGetBlipType(value)?.ToString() ?? $"BlipType:0x{value:X2}";

    internal static string? GetBlipRecordTypeName(ushort? recordType) => recordType switch {
        null => null,
        0xF01A => "OfficeArtBlipEMF",
        0xF01B => "OfficeArtBlipWMF",
        0xF01C => "OfficeArtBlipPICT",
        0xF01D => "OfficeArtBlipJPEG",
        0xF01E => "OfficeArtBlipPNG",
        0xF01F => "OfficeArtBlipDIB",
        0xF029 => "OfficeArtBlipTIFF",
        0xF02A => "OfficeArtBlipJPEG",
        _ => $"BlipRecordType:0x{recordType.Value:X4}"
    };

    internal static string? GetContentType(ushort? recordType, OfficeArtBlipType? recordInstanceType,
        OfficeArtBlipType? win32Type, OfficeArtBlipType? macOsType) {
        OfficeArtBlipType? type = GetBlipTypeFromRecord(recordType)
            ?? recordInstanceType ?? win32Type ?? macOsType;
        return type switch {
            OfficeArtBlipType.Png => "image/png",
            OfficeArtBlipType.Jpeg or OfficeArtBlipType.CmykJpeg => "image/jpeg",
            OfficeArtBlipType.Dib => "image/bmp",
            OfficeArtBlipType.Tiff => "image/tiff",
            OfficeArtBlipType.Emf => "image/x-emf",
            OfficeArtBlipType.Wmf => "image/x-wmf",
            _ => null
        };
    }

    private static OfficeArtBlipType? GetBlipTypeFromRecord(ushort? recordType) => recordType switch {
        0xF01A => OfficeArtBlipType.Emf,
        0xF01B => OfficeArtBlipType.Wmf,
        0xF01C => OfficeArtBlipType.Pict,
        0xF01D => OfficeArtBlipType.Jpeg,
        0xF01E => OfficeArtBlipType.Png,
        0xF01F => OfficeArtBlipType.Dib,
        0xF029 => OfficeArtBlipType.Tiff,
        0xF02A => OfficeArtBlipType.CmykJpeg,
        _ => null
    };
}
