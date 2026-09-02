using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeJpegCodec {
    /// <summary>Encodes an RGBA image directly to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficeJpegEncodeOptions? options = null) {
        EncodeTo(image, destination, options, CancellationToken.None);
    }

    internal static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficeJpegEncodeOptions? options,
        CancellationToken cancellationToken) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterOutput.EnsureWritable(destination);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeJpegEncodeOptions effectiveOptions = options ?? new OfficeJpegEncodeOptions();
        byte[] rgba = image.PixelBuffer;
        if (HasTransparency(rgba, cancellationToken)) {
            effectiveOptions = effectiveOptions.WithAdditionalRetainedManagedBytes(rgba.LongLength + 24L);
            rgba = (byte[])rgba.Clone();
            FlattenAlpha(rgba, effectiveOptions.Background, cancellationToken);
        }
        OfficeJpegWriter.WriteRgba(
            destination,
            image.Width,
            image.Height,
            rgba,
            checked(image.Width * 4),
            effectiveOptions,
            cancellationToken);
    }

#if NET8_0_OR_GREATER
    /// <summary>Encodes an RGBA image directly to a caller-owned buffer writer.</summary>
    public static void EncodeTo(
        OfficeRasterImage image,
        IBufferWriter<byte> destination,
        OfficeJpegEncodeOptions? options = null) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        EncodeTo(image, stream, options);
    }
#endif
}
