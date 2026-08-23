using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    /// <summary>Encodes a single RGBA image directly to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficeTiffEncodeOptions? options = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterOutput.EnsureWritable(destination);
        OfficeTiffEncodeOptions effective = options ?? new OfficeTiffEncodeOptions();
        ValidateOptions(effective);

        byte[] pixels = image.PixelBuffer;
        byte[] compressionInput = PrepareTiffCompressionInput(pixels, image.Width, image.Height, effective);
        byte[]? compressed = effective.Compression switch {
            OfficeTiffCompression.Deflate => OfficeZlibCodec.Compress(compressionInput),
            OfficeTiffCompression.Lzw => EncodeLzw(compressionInput),
            _ => null
        };
        int stripLength = effective.Compression switch {
            OfficeTiffCompression.None => pixels.Length,
            OfficeTiffCompression.Lzw => compressed!.Length,
            OfficeTiffCompression.PackBits => EncodePackBits(pixels, output: null, outputOffset: 0),
            OfficeTiffCompression.Deflate => compressed!.Length,
            _ => throw new ArgumentOutOfRangeException(nameof(options))
        };

        byte[] header = CreateEncodingHeader(image, effective, stripLength);
        destination.Write(header, 0, header.Length);
        switch (effective.Compression) {
            case OfficeTiffCompression.None:
                destination.Write(pixels, 0, pixels.Length);
                break;
            case OfficeTiffCompression.PackBits:
                int written = WritePackBits(pixels, destination);
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
        int entryCount = BaseEntryCount + (writePredictor ? 1 : 0);
        int ifdLength = 2 + (entryCount * 12) + 4;
        int bitsPerSampleOffset = checked(ifdOffset + ifdLength);
        int xResolutionOffset = checked(bitsPerSampleOffset + 8);
        int yResolutionOffset = checked(xResolutionOffset + 8);
        int stripOffset = checked(yResolutionOffset + 8);
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
        WriteEntry(output, ref entry, 282, 5, 1, xResolutionOffset);
        WriteEntry(output, ref entry, 283, 5, 1, yResolutionOffset);
        WriteShortEntry(output, ref entry, 284, 1);
        WriteShortEntry(output, ref entry, 296, 2);
        if (writePredictor) WriteShortEntry(output, ref entry, 317, (int)options.Predictor);
        WriteShortEntry(output, ref entry, 338, 2);
        WriteUInt32(output, entry, 0);

        WriteUInt16(output, bitsPerSampleOffset, 8);
        WriteUInt16(output, bitsPerSampleOffset + 2, 8);
        WriteUInt16(output, bitsPerSampleOffset + 4, 8);
        WriteUInt16(output, bitsPerSampleOffset + 6, 8);
        WriteRational(output, xResolutionOffset, options.DpiX);
        WriteRational(output, yResolutionOffset, options.DpiY);
        return output;
    }

    private static int WritePackBits(byte[] input, Stream destination) {
        int index = 0;
        int written = 0;
        while (index < input.Length) {
            int runLength = CountRun(input, index);
            if (runLength >= 3) {
                destination.WriteByte(unchecked((byte)(257 - runLength)));
                destination.WriteByte(input[index]);
                written = checked(written + 2);
                index += runLength;
                continue;
            }

            int literalStart = index;
            int literalLength = 0;
            while (index < input.Length && literalLength < 128) {
                runLength = CountRun(input, index);
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
