using System;
using System.IO;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    /// <summary>Encodes an image with an explicit lossless or lossy WebP compression mode.</summary>
    /// <remarks>Lossy mode preserves alpha exactly and supports dimensions from 1 through 16,383 pixels.</remarks>
    public static byte[] Encode(OfficeRasterImage image, OfficeWebpEncodeOptions options) {
        ValidateWebpOptions(options);
        if (options.Mode == OfficeWebpEncodingMode.Lossless && options.RetainedManagedBytes == 0L) {
            return options.WritePhysicalResolution
                ? Encode(image, options.DpiX, options.DpiY) : Encode(image);
        }
        return Encode(image, options, CancellationToken.None);
    }

    /// <summary>Encodes an image with cancellation and an explicit WebP compression mode.</summary>
    public static byte[] Encode(OfficeRasterImage image, OfficeWebpEncodeOptions options, CancellationToken cancellationToken) {
        using var destination = new MemoryStream();
        ValidateWebpOptions(options);
        EncodeTo(image, destination, options,
            options.WritePhysicalResolution ? options.DpiX : (double?)null,
            options.WritePhysicalResolution ? options.DpiY : (double?)null,
            cancellationToken, materializeOutput: true);
        cancellationToken.ThrowIfCancellationRequested();
        return destination.ToArray();
    }

    /// <summary>Encodes WebP to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open. Lossy mode buffers bounded VP8 partitions before writing its RIFF container.</remarks>
    public static void EncodeTo(OfficeRasterImage image, Stream destination, OfficeWebpEncodeOptions options,
        CancellationToken cancellationToken = default) {
        ValidateWebpOptions(options);
        EncodeTo(image, destination, options,
            options.WritePhysicalResolution ? options.DpiX : (double?)null,
            options.WritePhysicalResolution ? options.DpiY : (double?)null,
            cancellationToken);
    }

#if NET8_0_OR_GREATER
    /// <summary>Encodes WebP to a caller-owned buffer writer.</summary>
    public static void EncodeTo(OfficeRasterImage image, IBufferWriter<byte> destination, OfficeWebpEncodeOptions options,
        CancellationToken cancellationToken = default) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        EncodeTo(image, stream, options, cancellationToken);
    }
#endif

    internal static void EncodeTo(OfficeRasterImage image, Stream destination, OfficeWebpEncodeOptions options,
        double? dpiX, double? dpiY, CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null, bool materializeOutput = false) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        ValidateWebpOptions(options);
        OfficeRasterOutput.EnsureWritable(destination);
        cancellationToken.ThrowIfCancellationRequested();
        if (options.Mode == OfficeWebpEncodingMode.Lossless) {
            bool resolution = dpiX.HasValue && dpiY.HasValue;
            if (resolution) {
                ValidateDpi(dpiX!.Value, nameof(dpiX));
                ValidateDpi(dpiY!.Value, nameof(dpiY));
            }
            EncodeStreaming(image, destination, resolution, dpiX ?? 96D, dpiY ?? 96D,
                cancellationToken, checkpointObserver, options.RetainedManagedBytes, materializeOutput);
            return;
        }
        if (image.Width > 16383 || image.Height > 16383) {
            throw new ArgumentOutOfRangeException(nameof(image), "Lossy VP8 dimensions cannot exceed 16,383 pixels.");
        }
        bool writeResolution = dpiX.HasValue && dpiY.HasValue;
        if (writeResolution) {
            ValidateDpi(dpiX!.Value, nameof(dpiX));
            ValidateDpi(dpiY!.Value, nameof(dpiY));
        }
        byte[] pixels = image.PixelBuffer;
        bool hasAlpha = HasTransparency(pixels, cancellationToken);
        long paddedPixels = checked(((image.Width + 15L) / 16L * 16L) * ((image.Height + 15L) / 16L * 16L));
        // Each macroblock has at most 384 coefficient tokens, bounded literal magnitudes,
        // and mode headers. The boolean writers enforce this conservative per-image ceiling.
        int payloadCeiling = (int)Math.Min(OfficeRasterGuards.MaximumEncodedBytes, checked(paddedPixels * 8L + 4096L));
        long alphaBytes = hasAlpha ? (long)image.Width * image.Height + 1L : 0L;
        long resolutionBytes = writeResolution ? 128L : 0L;
        OfficeRasterOutput.TryGetMemoryStream(destination, out MemoryStream? outputStream);
        long retainedOutputBytes = outputStream == null ? 0L
            : checked(OfficeRasterOutput.GetMemoryStreamBackingBytes(outputStream) + 24L);
        // Reserve source/padded/reconstructed planes, raw alpha and block scratch up front.
        // Entropy buffers account for growth as it occurs, and the completed file determines
        // the MemoryStream growth/final-copy reserve before any container bytes are written.
        long workingBytes = checked(options.RetainedManagedBytes + pixels.LongLength + paddedPixels * 5L +
            alphaBytes * 2L + 128L * 1024L + retainedOutputBytes);
        var memory = new OfficeVp8EncodingMemory(workingBytes);
        OfficeVp8Encoder.EncodeRgba32(pixels, image.Width, image.Height, options.Quality,
            payloadCeiling, memory, cancellationToken, checkpointObserver, out byte[] vp8, out byte[] alpha);
        byte[]? exif = writeResolution ? CreateResolutionExif(image.Width, image.Height, dpiX!.Value, dpiY!.Value) : null;
        int length = checked(12 + 8 + vp8.Length + (vp8.Length & 1) +
            ((alpha.Length > 0 || exif != null) ? 18 : 0) +
            (alpha.Length > 0 ? 8 + alpha.Length + (alpha.Length & 1) : 0) +
            (exif != null ? 8 + exif.Length + (exif.Length & 1) : 0));
        if (length > OfficeRasterGuards.MaximumEncodedBytes) {
            throw new ArgumentException("WebP output exceeds encoded-size limits.", nameof(image));
        }
        long outputPeakBytes = outputStream == null ? 0L
            : OfficeRasterOutput.GetMemoryStreamWritePeakBytes(outputStream, length, materializeOutput);
        memory.EnsureAdditionalBytes(checked(resolutionBytes + outputPeakBytes - retainedOutputBytes));
        var header = new byte[12];
        WriteAscii(header, 0, "RIFF");
        WriteUInt32(header, 4, length - 8);
        WriteAscii(header, 8, "WEBP");
        destination.Write(header, 0, header.Length);
        if (alpha.Length > 0 || exif != null) {
            var vp8x = new byte[10];
            vp8x[0] = (byte)((alpha.Length > 0 ? 0x10 : 0) | (exif != null ? 0x08 : 0));
            WriteUInt24(vp8x, 4, image.Width - 1);
            WriteUInt24(vp8x, 7, image.Height - 1);
            WriteLossyWebpChunk(destination, "VP8X", vp8x, cancellationToken);
        }
        if (alpha.Length > 0) WriteLossyWebpChunk(destination, "ALPH", alpha, cancellationToken);
        WriteLossyWebpChunk(destination, "VP8 ", vp8, cancellationToken);
        if (exif != null) WriteLossyWebpChunk(destination, "EXIF", exif, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static void ValidateWebpOptions(OfficeWebpEncodeOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.Mode != OfficeWebpEncodingMode.Lossless && options.Mode != OfficeWebpEncodingMode.Lossy) {
            throw new ArgumentOutOfRangeException(nameof(options), "Unknown WebP compression mode.");
        }
        if (options.Mode == OfficeWebpEncodingMode.Lossy && (options.Quality < 1 || options.Quality > 100)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Lossy WebP quality must be from 1 through 100.");
        }
        if (options.RetainedManagedBytes < 0) throw new ArgumentOutOfRangeException(nameof(options));
    }

    private static void WriteLossyWebpChunk(Stream destination, string name, byte[] payload, CancellationToken cancellationToken) {
        var header = new byte[8];
        WriteAscii(header, 0, name);
        WriteUInt32(header, 4, payload.Length);
        destination.Write(header, 0, header.Length);
        for (int offset = 0; offset < payload.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(16384, payload.Length - offset);
            destination.Write(payload, offset, count);
            offset += count;
        }
        if ((payload.Length & 1) != 0) destination.WriteByte(0);
    }
}
