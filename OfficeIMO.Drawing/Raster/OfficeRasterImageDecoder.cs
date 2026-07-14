namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free decoder for raster image bytes that can be painted by <see cref="OfficeRasterCanvas"/>.
/// </summary>
public static class OfficeRasterImageDecoder {
    /// <summary>
    /// Human-readable summary of raster formats currently decoded by the managed renderer.
    /// </summary>
    public const string SupportedFormatDescription = "PNG, JPEG, uncompressed BMP, and first-frame GIF image bytes";

    /// <summary>
    /// Attempts to decode image bytes into an RGBA raster buffer supported by dependency-free export.
    /// </summary>
    public static bool TryDecode(byte[]? bytes, out OfficeRasterImage? image) =>
        OfficePngReader.TryDecode(bytes!, out image) ||
        OfficeJpegCodec.TryDecode(bytes, out image) ||
        OfficeBmpReader.TryDecode(bytes, out image) ||
        OfficeGifReader.TryDecode(bytes, out image);
}
