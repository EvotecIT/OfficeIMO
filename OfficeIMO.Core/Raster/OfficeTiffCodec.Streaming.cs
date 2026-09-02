using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    /// <summary>Encodes a single RGBA image directly to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficeTiffEncodeOptions? options = null) {
        EncodeTo(image, destination, options, CancellationToken.None);
    }

    internal static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficeTiffEncodeOptions? options,
        CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterOutput.EnsureWritable(destination);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeTiffEncodeOptions effective = options ?? new OfficeTiffEncodeOptions();
        ValidateOptions(effective);

        byte[] pixels = image.PixelBuffer;
        EnsureSinglePageCompressionWorkingSet(pixels.LongLength, effective);
        byte[]? compressed = effective.Compression switch {
            OfficeTiffCompression.Deflate => OfficeZlibCodec.Compress(
                PrepareTiffCompressionInput(pixels, image.Width, image.Height, effective, cancellationToken, checkpointObserver),
                cancellationToken),
            OfficeTiffCompression.Lzw => EncodeTiffLzw(pixels, image.Width, image.Height, effective, cancellationToken, checkpointObserver),
            _ => null
        };
        int stripLength = effective.Compression switch {
            OfficeTiffCompression.None => pixels.Length,
            OfficeTiffCompression.Lzw => compressed!.Length,
            OfficeTiffCompression.PackBits =>
                EncodePackBitsRows(pixels, image.Width * 4, image.Height, output: null, outputOffset: 0, cancellationToken, checkpointObserver),
            OfficeTiffCompression.Deflate => compressed!.Length,
            _ => throw new ArgumentOutOfRangeException(nameof(options))
        };

        byte[] header = CreateEncodingHeader(image, effective, stripLength);
        OfficeRasterGuards.EnsureOutputBytes(
            checked((long)header.Length + stripLength),
            "The TIFF exceeds the encoded-size limit.");
        long retainedStripBytes = effective.Compression is OfficeTiffCompression.Lzw or OfficeTiffCompression.Deflate
            ? stripLength
            : 0L;
        long retainedOutputBytes = OfficeRasterOutput.TryGetMemoryStream(destination, out _)
            ? checked(2L * (header.LongLength + stripLength))
            : header.LongLength;
        if (!IsMultiPageTiffWorkingSetWithinLimit(pixels.LongLength, retainedStripBytes, retainedOutputBytes)) {
            throw new ArgumentException("The TIFF encoding working set exceeds the managed limit.", nameof(image));
        }
        cancellationToken.ThrowIfCancellationRequested();
        destination.Write(header, 0, header.Length);
        switch (effective.Compression) {
            case OfficeTiffCompression.None:
                destination.Write(pixels, 0, pixels.Length);
                break;
            case OfficeTiffCompression.PackBits:
                int written = WritePackBitsRows(pixels, image.Width * 4, image.Height, destination, cancellationToken);
                if (written != stripLength) {
                    throw new InvalidOperationException("The TIFF PackBits length calculation is inconsistent.");
                }
                break;
            case OfficeTiffCompression.Lzw:
            case OfficeTiffCompression.Deflate:
                destination.Write(compressed!, 0, compressed!.Length);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(options));
        }
    }

#if NET8_0_OR_GREATER
    /// <summary>Encodes a single RGBA image directly to a caller-owned buffer writer.</summary>
    public static void EncodeTo(
        OfficeRasterImage image,
        IBufferWriter<byte> destination,
        OfficeTiffEncodeOptions? options = null) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        EncodeTo(image, stream, options);
    }
#endif

    private static byte[] CreateEncodingHeader(
        OfficeRasterImage image,
        OfficeTiffEncodeOptions options,
        int stripLength) {
        const int ifdOffset = 8;
        bool writePredictor = UsesHorizontalPredictor(options);
        int entryCount = BaseEntryCount - (options.WriteResolution ? 0 : 3) + (writePredictor ? 1 : 0);
        int ifdLength = 2 + (entryCount * 12) + 4;
        int bitsPerSampleOffset = checked(ifdOffset + ifdLength);
        int xResolutionOffset = checked(bitsPerSampleOffset + 8);
        int yResolutionOffset = checked(xResolutionOffset + 8);
        int stripOffset = options.WriteResolution ? checked(yResolutionOffset + 8) : xResolutionOffset;
        var output = new byte[stripOffset];

        output[0] = (byte)'I';
        output[1] = (byte)'I';
        WriteUInt16(output, 2, 42);
        WriteUInt32(output, 4, ifdOffset);
        WriteUInt16(output, ifdOffset, entryCount);

        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, image.Width);
        WriteEntry(output, ref entry, 257, 4, 1, image.Height);
        WriteEntry(output, ref entry, 258, 3, 4, bitsPerSampleOffset);
        WriteShortEntry(output, ref entry, 259, (int)options.Compression);
        WriteShortEntry(output, ref entry, 262, 2);
        WriteEntry(output, ref entry, 273, 4, 1, stripOffset);
        WriteShortEntry(output, ref entry, 274, 1);
        WriteShortEntry(output, ref entry, 277, 4);
        WriteEntry(output, ref entry, 278, 4, 1, image.Height);
        WriteEntry(output, ref entry, 279, 4, 1, stripLength);
        if (options.WriteResolution) {
            WriteEntry(output, ref entry, 282, 5, 1, xResolutionOffset);
            WriteEntry(output, ref entry, 283, 5, 1, yResolutionOffset);
        }
        WriteShortEntry(output, ref entry, 284, 1);
        if (options.WriteResolution) WriteShortEntry(output, ref entry, 296, 2);
        if (writePredictor) WriteShortEntry(output, ref entry, 317, (int)options.Predictor);
        WriteShortEntry(output, ref entry, 338, 2);
        WriteUInt32(output, entry, 0);

        WriteUInt16(output, bitsPerSampleOffset, 8);
        WriteUInt16(output, bitsPerSampleOffset + 2, 8);
        WriteUInt16(output, bitsPerSampleOffset + 4, 8);
        WriteUInt16(output, bitsPerSampleOffset + 6, 8);
        if (options.WriteResolution) {
            WriteRational(output, xResolutionOffset, options.DpiX);
            WriteRational(output, yResolutionOffset, options.DpiY);
        }
        return output;
    }

    private static int WritePackBitsRows(
        byte[] input,
        int rowBytes,
        int rowCount,
        Stream destination,
        CancellationToken cancellationToken) {
        if (rowBytes <= 0 || rowCount <= 0 || (long)rowBytes * rowCount != input.Length) {
            throw new ArgumentException("TIFF PackBits row dimensions do not match the input buffer.");
        }
        int written = 0;
        for (int row = 0; row < rowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            written = checked(written + WritePackBits(
                input, row * rowBytes, rowBytes, destination, cancellationToken));
        }
        return written;
    }

    private static int WritePackBits(
        byte[] input,
        int inputOffset,
        int inputCount,
        Stream destination,
        CancellationToken cancellationToken) {
        int index = inputOffset;
        int inputEnd = checked(inputOffset + inputCount);
        int written = 0;
        while (index < inputEnd) {
            if ((index & 0x3FFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            int runLength = CountRun(input, index, inputEnd);
            if (runLength >= 3) {
                destination.WriteByte(unchecked((byte)(257 - runLength)));
                destination.WriteByte(input[index]);
                written = checked(written + 2);
                index += runLength;
                continue;
            }

            int literalStart = index;
            int literalLength = 0;
            while (index < inputEnd && literalLength < 128) {
                runLength = CountRun(input, index, inputEnd);
                if (runLength >= 3) break;
                int take = Math.Min(runLength, 128 - literalLength);
                index += take;
                literalLength += take;
            }

            destination.WriteByte((byte)(literalLength - 1));
            destination.Write(input, literalStart, literalLength);
            written = checked(written + literalLength + 1);
        }
        return written;
    }
}
