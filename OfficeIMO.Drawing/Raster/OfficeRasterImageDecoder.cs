namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free decoder for raster image bytes that can be painted by <see cref="OfficeRasterCanvas"/>.
/// </summary>
public static class OfficeRasterImageDecoder {
    /// <summary>
    /// Attempts to decode image bytes into an RGBA raster buffer supported by dependency-free export.
    /// </summary>
    public static bool TryDecode(byte[]? bytes, out OfficeRasterImage? image) =>
        OfficePngReader.TryDecode(bytes!, out image) ||
        OfficeBmpReader.TryDecode(bytes, out image) ||
        OfficeGifReader.TryDecode(bytes, out image);
}
