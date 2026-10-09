using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>The TIFF directory containing an Exif field.</summary>
public enum OfficeExifDirectory {
    /// <summary>Primary image attributes.</summary>
    Image,
    /// <summary>Camera and capture attributes.</summary>
    Exif,
    /// <summary>Geographic coordinates and receiver attributes.</summary>
    Gps,
    /// <summary>Interoperability attributes.</summary>
    Interoperability
}

/// <summary>The on-disk TIFF field representation.</summary>
public enum OfficeExifDataType {
    /// <summary>Unsigned eight-bit integers.</summary>
    Byte = 1,
    /// <summary>NUL-terminated ASCII text.</summary>
    Ascii = 2,
    /// <summary>Unsigned sixteen-bit integers.</summary>
    Short = 3,
    /// <summary>Unsigned thirty-two-bit integers.</summary>
    Long = 4,
    /// <summary>Unsigned rational numbers.</summary>
    Rational = 5,
    /// <summary>Signed eight-bit integers.</summary>
    SignedByte = 6,
    /// <summary>Uninterpreted bytes.</summary>
    Undefined = 7,
    /// <summary>Signed sixteen-bit integers.</summary>
    SignedShort = 8,
    /// <summary>Signed thirty-two-bit integers.</summary>
    SignedLong = 9,
    /// <summary>Signed rational numbers.</summary>
    SignedRational = 10,
    /// <summary>IEEE single precision values.</summary>
    Float = 11,
    /// <summary>IEEE double precision values.</summary>
    Double = 12
}

/// <summary>A portable Exif tag identity, including its directory and expected value representation.</summary>
public readonly struct OfficeExifTag : IEquatable<OfficeExifTag> {
    /// <summary>Creates a standard or application-defined tag.</summary>
    public OfficeExifTag(ushort id, OfficeExifDataType dataType, OfficeExifDirectory directory = OfficeExifDirectory.Image, string? name = null) {
        if ((int)dataType < 1 || (int)dataType > 12) throw new ArgumentOutOfRangeException(nameof(dataType));
        if ((int)directory < 0 || (int)directory > 3) throw new ArgumentOutOfRangeException(nameof(directory));
        Id = id;
        DataType = dataType;
        Directory = directory;
        Name = name ?? "0x" + id.ToString("X4");
    }
    /// <summary>The unsigned TIFF tag identifier.</summary>
    public ushort Id { get; }
    /// <summary>The expected TIFF representation.</summary>
    public OfficeExifDataType DataType { get; }
    /// <summary>The directory in which the tag is stored.</summary>
    public OfficeExifDirectory Directory { get; }
    /// <summary>The descriptive name, or hexadecimal identifier for an unknown tag.</summary>
    public string Name { get; }
    /// <inheritdoc />
    public bool Equals(OfficeExifTag other) => Id == other.Id && Directory == other.Directory;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeExifTag tag && Equals(tag);
    /// <inheritdoc />
    public override int GetHashCode() => (int)Directory * 65536 + Id;
    /// <inheritdoc />
    public override string ToString() => Name ?? "0x" + Id.ToString("X4");
    /// <summary>Image width in pixels.</summary>
    public static OfficeExifTag ImageWidth => new OfficeExifTag(256, OfficeExifDataType.Long, name: nameof(ImageWidth));
    /// <summary>Image height in pixels.</summary>
    public static OfficeExifTag ImageLength => new OfficeExifTag(257, OfficeExifDataType.Long, name: nameof(ImageLength));
    /// <summary>Image description.</summary>
    public static OfficeExifTag ImageDescription => new OfficeExifTag(270, OfficeExifDataType.Ascii, name: nameof(ImageDescription));
    /// <summary>Camera manufacturer.</summary>
    public static OfficeExifTag Make => new OfficeExifTag(271, OfficeExifDataType.Ascii, name: nameof(Make));
    /// <summary>Camera model.</summary>
    public static OfficeExifTag Model => new OfficeExifTag(272, OfficeExifDataType.Ascii, name: nameof(Model));
    /// <summary>Display orientation, represented by an unsigned short.</summary>
    public static OfficeExifTag Orientation => new OfficeExifTag(274, OfficeExifDataType.Short, name: nameof(Orientation));
    /// <summary>Horizontal resolution.</summary>
    public static OfficeExifTag XResolution => new OfficeExifTag(282, OfficeExifDataType.Rational, name: nameof(XResolution));
    /// <summary>Vertical resolution.</summary>
    public static OfficeExifTag YResolution => new OfficeExifTag(283, OfficeExifDataType.Rational, name: nameof(YResolution));
    /// <summary>Resolution unit, where two means inches and three means centimeters.</summary>
    public static OfficeExifTag ResolutionUnit => new OfficeExifTag(296, OfficeExifDataType.Short, name: nameof(ResolutionUnit));
    /// <summary>Image-producing software.</summary>
    public static OfficeExifTag Software => new OfficeExifTag(305, OfficeExifDataType.Ascii, name: nameof(Software));
    /// <summary>Image modification date and time.</summary>
    public static OfficeExifTag DateTime => new OfficeExifTag(306, OfficeExifDataType.Ascii, name: nameof(DateTime));
    /// <summary>Artist or creator.</summary>
    public static OfficeExifTag Artist => new OfficeExifTag(315, OfficeExifDataType.Ascii, name: nameof(Artist));
    /// <summary>Copyright notice.</summary>
    public static OfficeExifTag Copyright => new OfficeExifTag(33432, OfficeExifDataType.Ascii, name: nameof(Copyright));
    /// <summary>Exposure duration in seconds.</summary>
    public static OfficeExifTag ExposureTime => new OfficeExifTag(33434, OfficeExifDataType.Rational, OfficeExifDirectory.Exif, nameof(ExposureTime));
    /// <summary>Lens f-number.</summary>
    public static OfficeExifTag FNumber => new OfficeExifTag(33437, OfficeExifDataType.Rational, OfficeExifDirectory.Exif, nameof(FNumber));
    /// <summary>Photographic sensitivity.</summary>
    public static OfficeExifTag ISOSpeedRatings => new OfficeExifTag(34855, OfficeExifDataType.Short, OfficeExifDirectory.Exif, nameof(ISOSpeedRatings));
    /// <summary>Exif specification version as four bytes.</summary>
    public static OfficeExifTag ExifVersion => new OfficeExifTag(36864, OfficeExifDataType.Undefined, OfficeExifDirectory.Exif, nameof(ExifVersion));
    /// <summary>Original capture date and time.</summary>
    public static OfficeExifTag DateTimeOriginal => new OfficeExifTag(36867, OfficeExifDataType.Ascii, OfficeExifDirectory.Exif, nameof(DateTimeOriginal));
    /// <summary>Digitization date and time.</summary>
    public static OfficeExifTag DateTimeDigitized => new OfficeExifTag(36868, OfficeExifDataType.Ascii, OfficeExifDirectory.Exif, nameof(DateTimeDigitized));
    /// <summary>User comment including its Exif character-set prefix.</summary>
    public static OfficeExifTag UserComment => new OfficeExifTag(37510, OfficeExifDataType.Undefined, OfficeExifDirectory.Exif, nameof(UserComment));
    /// <summary>North or south latitude reference.</summary>
    public static OfficeExifTag GPSLatitudeRef => new OfficeExifTag(1, OfficeExifDataType.Ascii, OfficeExifDirectory.Gps, nameof(GPSLatitudeRef));
    /// <summary>Latitude degrees, minutes, and seconds.</summary>
    public static OfficeExifTag GPSLatitude => new OfficeExifTag(2, OfficeExifDataType.Rational, OfficeExifDirectory.Gps, nameof(GPSLatitude));
    /// <summary>East or west longitude reference.</summary>
    public static OfficeExifTag GPSLongitudeRef => new OfficeExifTag(3, OfficeExifDataType.Ascii, OfficeExifDirectory.Gps, nameof(GPSLongitudeRef));
    /// <summary>Longitude degrees, minutes, and seconds.</summary>
    public static OfficeExifTag GPSLongitude => new OfficeExifTag(4, OfficeExifDataType.Rational, OfficeExifDirectory.Gps, nameof(GPSLongitude));
}

/// <summary>An exact unsigned rational number used by Exif.</summary>
public readonly struct OfficeRational {
    /// <summary>Creates a rational, retaining a zero denominator when present in source metadata.</summary>
    public OfficeRational(uint numerator, uint denominator) { Numerator = numerator; Denominator = denominator; }
    /// <summary>The numerator.</summary>
    public uint Numerator { get; }
    /// <summary>The denominator.</summary>
    public uint Denominator { get; }
    /// <summary>Returns the quotient, or NaN for a zero denominator.</summary>
    public double ToDouble() => Denominator == 0 ? double.NaN : (double)Numerator / Denominator;
    /// <inheritdoc />
    public override string ToString() => Numerator + "/" + Denominator;
}

/// <summary>An exact signed rational number used by Exif.</summary>
public readonly struct OfficeSignedRational {
    /// <summary>Creates a signed rational value.</summary>
    public OfficeSignedRational(int numerator, int denominator) { Numerator = numerator; Denominator = denominator; }
    /// <summary>The numerator.</summary>
    public int Numerator { get; }
    /// <summary>The denominator.</summary>
    public int Denominator { get; }
    /// <summary>Returns the quotient, or NaN for a zero denominator.</summary>
    public double ToDouble() => Denominator == 0 ? double.NaN : (double)Numerator / Denominator;
    /// <inheritdoc />
    public override string ToString() => Numerator + "/" + Denominator;
}

/// <summary>An immutable typed Exif field snapshot.</summary>
public sealed class OfficeExifValue {
    private readonly object _value;
    private readonly byte[]? _parsedEncoding;
    private readonly int _parsedOffset;
    private readonly int _parsedLength;
    private readonly uint _parsedCount;
    private readonly bool _parsedLittleEndian;
    internal OfficeExifValue(OfficeExifTag tag, object value) { Tag = tag; _value = value is Array array ? array.Clone() : value; }
    /// <summary>Owns a validated parsed field without applying the stricter rules for new user edits.</summary>
    internal OfficeExifValue(OfficeExifProfileCodec.Field field, byte[] source, bool littleEndian, bool copyEncoding = true) : this(field.Tag, field.Value) {
        _parsedLength = field.ValueLength;
        if (copyEncoding) {
            _parsedEncoding = new byte[field.ValueLength];
            Buffer.BlockCopy(source, field.ValueOffset, _parsedEncoding, 0, _parsedLength);
        } else {
            // Classic Exif profiles already own immutable bounded bytes. Borrow those
            // bytes instead of allocating another payload for each value enumeration.
            _parsedEncoding = source;
            _parsedOffset = field.ValueOffset;
        }
        _parsedCount = field.Count;
        _parsedLittleEndian = littleEndian;
    }
    /// <summary>The field identity.</summary>
    public OfficeExifTag Tag { get; }
    /// <summary>The TIFF representation.</summary>
    public OfficeExifDataType DataType => Tag.DataType;
    /// <summary>The scalar, text, or typed array value. Arrays are copied.</summary>
    public object Value => _value is Array array ? array.Clone() : _value;
    internal long EncodedByteLength => _parsedEncoding != null ? _parsedLength : (_value is string text ? text.Length + 1L : _value is Array array ? array.LongLength * OfficeExifProfileCodec.Size((int)DataType) : OfficeExifProfileCodec.Size((int)DataType));
    internal long RetainedByteLength => checked(EncodedByteLength * 2L + (_parsedEncoding?.LongLength ?? 0L));
    /// <summary>Preserves parsed counts, ASCII string lists, and numeric bits while adapting byte order when a field is relocated.</summary>
    internal byte[] EncodeValue(bool littleEndian, out uint count, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (_parsedEncoding == null) return OfficeExifProfileCodec.EncodeValue(DataType, _value, littleEndian, out count, token);
        count = _parsedCount;
        byte[] encoded = new byte[_parsedLength];
        Buffer.BlockCopy(_parsedEncoding, _parsedOffset, encoded, 0, encoded.Length);
        int wordBytes = DataType == OfficeExifDataType.Rational || DataType == OfficeExifDataType.SignedRational ? 4 : OfficeExifProfileCodec.Size((int)DataType);
        if (_parsedLittleEndian != littleEndian && wordBytes > 1) {
            for (int at = 0; at < encoded.Length; at += wordBytes) {
                if ((at & 4095) == 0) token.ThrowIfCancellationRequested();
                Array.Reverse(encoded, at, wordBytes);
            }
        }
        token.ThrowIfCancellationRequested();
        return encoded;
    }
}
