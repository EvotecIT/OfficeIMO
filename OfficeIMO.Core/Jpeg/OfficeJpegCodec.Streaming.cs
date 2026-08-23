using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;

namespace OfficeIMO.Drawing;

public static partial class OfficeJpegCodec {
    /// <summary>Encodes an RGBA image directly to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void Encode(
        OfficeRasterImage image,
        Stream destination,
        OfficeJpegEncodeOptions? options = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterOutput.EnsureWritable(destination);
        OfficeJpegEncodeOptions effectiveOptions = options ?? new OfficeJpegEncodeOptions();
        byte[] rgba = image.PixelBuffer;
        if (HasTransparency(rgba)) {
            rgba = (byte[])rgba.Clone();
            FlattenAlpha(rgba, effectiveOptions.Background);
        }
        OfficeJpegWriter.WriteRgba(
            destination,
            image.Width,
            image.Height,
            rgba,
            checked(image.Width * 4),
            effectiveOptions);
    }

#if NET8_0_OR_GREATER
    /// <summary>Encodes an RGBA image directly to a caller-owned buffer writer.</summary>
    public static void Encode(
        OfficeRasterImage image,
        IBufferWriter<byte> destination,
        OfficeJpegEncodeOptions? options = null) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        Encode(image, stream, options);
    }
#endif
}
