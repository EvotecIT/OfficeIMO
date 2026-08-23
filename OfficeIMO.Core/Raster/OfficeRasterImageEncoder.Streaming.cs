using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterImageEncoder {
    /// <summary>Encodes an RGBA image directly to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        Stream destination,
        OfficeRasterEncodingOptions? options = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterOutput.EnsureWritable(destination);
        OfficeRasterEncodingOptions effective =
            (options ?? new OfficeRasterEncodingOptions()).Resolve(format);

        switch (format) {
            case OfficeImageExportFormat.Png:
                OfficePngWriter.EncodeTo(
                    image,
                    destination,
                    effective.Png ?? throw new InvalidOperationException("PNG encoding options cannot be null."));
                break;
            case OfficeImageExportFormat.Jpeg:
                OfficeJpegCodec.EncodeTo(
                    image,
                    destination,
                    effective.Jpeg ?? throw new InvalidOperationException("JPEG encoding options cannot be null."));
                break;
            case OfficeImageExportFormat.Tiff:
                OfficeTiffCodec.EncodeTo(
                    image,
                    destination,
                    effective.Tiff ?? throw new InvalidOperationException("TIFF encoding options cannot be null."));
                break;
            case OfficeImageExportFormat.Webp:
                OfficeWebpCodec.EncodeTo(image, destination, effective.DpiX, effective.DpiY);
                break;
            case OfficeImageExportFormat.Svg:
                throw new ArgumentException("SVG output requires a vector renderer.", nameof(format));
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
    }

#if NET8_0_OR_GREATER
    /// <summary>Encodes an RGBA image directly to a caller-owned buffer writer.</summary>
    public static void EncodeTo(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        IBufferWriter<byte> destination,
        OfficeRasterEncodingOptions? options = null) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        EncodeTo(image, format, stream, options);
    }
#endif
}
