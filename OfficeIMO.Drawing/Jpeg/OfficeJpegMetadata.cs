namespace OfficeIMO.Drawing;

/// <summary>
/// Optional JPEG metadata payloads (EXIF/XMP/ICC).
/// </summary>
public readonly struct OfficeJpegMetadata {
    private readonly byte[]? _exif;
    private readonly byte[]? _xmp;
    private readonly byte[]? _icc;

    /// <summary>
    /// EXIF payload (TIFF data, optional "Exif\0\0" header).
    /// </summary>
    public byte[]? Exif => Clone(_exif);

    /// <summary>
    /// XMP payload (RDF/XML, optional XMP namespace header).
    /// </summary>
    public byte[]? Xmp => Clone(_xmp);

    /// <summary>
    /// ICC profile payload.
    /// </summary>
    public byte[]? Icc => Clone(_icc);

    /// <summary>
    /// Indicates whether any metadata is present.
    /// </summary>
    public bool HasData => (_exif is { Length: > 0 }) || (_xmp is { Length: > 0 }) || (_icc is { Length: > 0 });

    /// <summary>
    /// Creates metadata with optional EXIF/XMP/ICC payloads.
    /// </summary>
    public OfficeJpegMetadata(byte[]? exif = null, byte[]? xmp = null, byte[]? icc = null) {
        _exif = Clone(exif);
        _xmp = Clone(xmp);
        _icc = Clone(icc);
    }

    private static byte[]? Clone(byte[]? value) => value == null ? null : (byte[])value.Clone();
}
