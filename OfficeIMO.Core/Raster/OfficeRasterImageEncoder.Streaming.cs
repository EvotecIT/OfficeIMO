using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal enum OfficeRasterEncodingCheckpoint {
    JpegCoefficientRow,
    TiffCompressionRow
}

public static partial class OfficeRasterImageEncoder {
    /// <summary>Encodes an RGBA image directly to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        Stream destination,
        OfficeRasterEncodingOptions? options = null) {
        EncodeToCore(image, format, destination, options, CancellationToken.None);
    }

    /// <summary>
    /// Encodes an RGBA image directly to a caller-owned writable stream while observing cancellation
    /// and enforcing an encoded-byte ceiling as bytes are produced.
    /// </summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        Stream destination,
        OfficeRasterEncodingOptions? options,
        long maximumEncodedBytes,
        CancellationToken cancellationToken = default) {
        var budget = new OfficeImageExportEncodingBudget(maximumEncodedBytes);
        EncodeTo(image, format, destination, options, budget, cancellationToken);
    }

    internal static void EncodeTo(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        Stream destination,
        OfficeRasterEncodingOptions? options,
        OfficeImageExportEncodingBudget budget,
        CancellationToken cancellationToken) {
        if (budget == null) throw new ArgumentNullException(nameof(budget));
        using var guarded = new OfficeImageExportEncodingStream(destination, budget, cancellationToken);
        EncodeToCore(image, format, guarded, options, cancellationToken, checkpointObserver: null);
        cancellationToken.ThrowIfCancellationRequested();
    }

    internal static void EncodeTo(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        Stream destination,
        OfficeRasterEncodingOptions? options,
        long maximumEncodedBytes,
        CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint> checkpointObserver) {
        if (checkpointObserver == null) throw new ArgumentNullException(nameof(checkpointObserver));
        var budget = new OfficeImageExportEncodingBudget(maximumEncodedBytes);
        using var guarded = new OfficeImageExportEncodingStream(destination, budget, cancellationToken);
        EncodeToCore(image, format, guarded, options, cancellationToken, checkpointObserver);
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static void EncodeToCore(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        Stream destination,
        OfficeRasterEncodingOptions? options,
        CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterOutput.EnsureWritable(destination);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterEncodingOptions effective =
            (options ?? new OfficeRasterEncodingOptions()).Resolve(format);

        switch (format) {
            case OfficeImageExportFormat.Png:
                OfficePngWriter.EncodeTo(
                    image,
                    destination,
                    effective.Png ?? throw new InvalidOperationException("PNG encoding options cannot be null."),
                    cancellationToken);
                break;
            case OfficeImageExportFormat.Jpeg:
                OfficeJpegCodec.EncodeTo(
                    image,
                    destination,
                    effective.Jpeg ?? throw new InvalidOperationException("JPEG encoding options cannot be null."),
                    cancellationToken,
                    checkpointObserver);
                break;
            case OfficeImageExportFormat.Tiff:
                OfficeTiffCodec.EncodeTo(
                    image,
                    destination,
                    effective.Tiff ?? throw new InvalidOperationException("TIFF encoding options cannot be null."),
                    cancellationToken,
                    checkpointObserver);
                break;
            case OfficeImageExportFormat.Webp:
                OfficeWebpCodec.EncodeTo(
                    image,
                    destination,
                    effective.WriteResolutionMetadata ? effective.DpiX : (double?)null,
                    effective.WriteResolutionMetadata ? effective.DpiY : (double?)null,
                    cancellationToken);
                break;
            case OfficeImageExportFormat.Svg:
                throw new ArgumentException("SVG output requires a vector renderer.", nameof(format));
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
        cancellationToken.ThrowIfCancellationRequested();
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
