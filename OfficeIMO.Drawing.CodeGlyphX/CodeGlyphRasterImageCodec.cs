using System;
using System.IO;
using CodeGlyphX;
using CodeGlyphX.Rendering;

namespace OfficeIMO.Drawing.CodeGlyphX;

/// <summary>Optional managed CodeGlyphX raster decoder for Drawing and HTML/PDF exports.</summary>
/// <remarks>The shared Drawing decoder applies the request's container and pixel limits before
/// invoking this provider. Direct calls also retain CodeGlyphX's standard encoded-byte and pixel bounds.
/// The synchronous provider call is not preempted; the export observes cancellation around it.</remarks>
public sealed class CodeGlyphRasterImageCodec : IOfficeRasterImageCodec {
    /// <inheritdoc />
    public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
        image = null;
        if (encodedBytes == null || encodedBytes.Length == 0 || encodedBytes.Length > ImageReader.DefaultMaxImageBytes) return false;
        try {
            byte[] pixels = ImageReader.DecodeRgba32(encodedBytes, new ImageDecodeOptions {
                MaxPixels = ImageReader.DefaultMaxPixels,
                MaxBytes = ImageReader.DefaultMaxImageBytes
            }, out int width, out int height);
            image = OfficeRasterImage.FromOwnedRgba32(width, height, pixels);
            return true;
        } catch (Exception exception) when (exception is ArgumentException || exception is FormatException ||
            exception is NotSupportedException || exception is IOException || exception is OverflowException) {
            return false;
        }
    }
}
