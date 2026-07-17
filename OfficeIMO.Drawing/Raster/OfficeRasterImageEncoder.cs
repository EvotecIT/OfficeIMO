using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free encoder for raster export formats.
/// </summary>
public static class OfficeRasterImageEncoder {
    internal const int JpegMaximumDimension = ushort.MaxValue;
    internal const int WebpMaximumDimension = 16384;

    /// <summary>Returns the maximum supported pixel width or height for a raster format.</summary>
    public static int GetMaximumDimension(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => int.MaxValue,
        OfficeImageExportFormat.Jpeg => JpegMaximumDimension,
        OfficeImageExportFormat.Tiff => int.MaxValue,
        OfficeImageExportFormat.Webp => WebpMaximumDimension,
        OfficeImageExportFormat.Svg => throw new ArgumentException("SVG output does not have a raster dimension limit.", nameof(format)),
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

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
