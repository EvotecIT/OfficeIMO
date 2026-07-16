using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free encoder for raster export formats.
/// </summary>
public static class OfficeRasterImageEncoder {
    /// <summary>Encodes an RGBA image using the requested raster format.</summary>
    public static byte[] Encode(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        OfficeRasterEncodingOptions? options = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterEncodingOptions effective = options ?? new OfficeRasterEncodingOptions();
        return format switch {
            OfficeImageExportFormat.Png => OfficePngWriter.Encode(image),
            OfficeImageExportFormat.Jpeg => OfficeJpegCodec.Encode(
                image,
                effective.Jpeg ?? throw new InvalidOperationException("JPEG encoding options cannot be null.")),
            OfficeImageExportFormat.Tiff => OfficeTiffCodec.Encode(
                image,
                effective.Tiff ?? throw new InvalidOperationException("TIFF encoding options cannot be null.")),
            OfficeImageExportFormat.Webp => OfficeWebpCodec.Encode(image),
            OfficeImageExportFormat.Svg => throw new ArgumentException("SVG output requires a vector renderer.", nameof(format)),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };
    }
}
