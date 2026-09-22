using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    internal static bool TryValidateStripPayload(
        byte[] input,
        int inputOffset,
        int inputCount,
        int compression,
        int expectedCount,
        CancellationToken cancellationToken = default) {
        if (input == null || inputOffset < 0 || inputCount < 0 || expectedCount <= 0 ||
            inputOffset > input.Length - inputCount) return false;
        cancellationToken.ThrowIfCancellationRequested();
        return compression switch {
            (int)OfficeTiffCompression.None => inputCount == expectedCount,
            (int)OfficeTiffCompression.PackBits => TryValidatePackBitsPayload(
                input, inputOffset, inputCount, expectedCount, cancellationToken),
            (int)OfficeTiffCompression.Lzw => TryValidateLzwPayload(
                input, inputOffset, inputCount, expectedCount, cancellationToken),
            (int)OfficeTiffCompression.Deflate => OfficeZlibCodec.TryValidateExact(
                input, inputOffset, inputCount, expectedCount, cancellationToken),
            32946 => OfficeZlibCodec.TryValidateExact(
                    input, inputOffset, inputCount, expectedCount, cancellationToken) ||
                TryValidateRawDeflatePayload(input, inputOffset, inputCount, expectedCount, cancellationToken),
            _ => false
        };
    }

    private static bool TryValidatePackBitsPayload(
        byte[] input,
        int offset,
        int count,
        int expectedCount,
        CancellationToken cancellationToken) {
        int end = offset + count;
        int source = offset;
        int decoded = 0;
        while (source < end && decoded < expectedCount) {
            if (((source - offset) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int header = unchecked((sbyte)input[source++]);
            if (header >= 0) {
                int literalCount = header + 1;
                if (source > end - literalCount || decoded > expectedCount - literalCount) return false;
                source += literalCount;
                decoded += literalCount;
            } else if (header >= -127) {
                int repeatCount = 1 - header;
                if (source >= end || decoded > expectedCount - repeatCount) return false;
                source++;
                decoded += repeatCount;
            }
        }
        return decoded == expectedCount &&
            TryValidatePackBitsPadding(input, source, end - source, cancellationToken);
    }

    private static bool TryValidateRawDeflatePayload(
        byte[] input,
        int offset,
        int count,
        int expectedCount,
        CancellationToken cancellationToken) {
        if (!OfficeDeflateStreamValidator.TryValidateExact(
                input, offset, count, expectedCount, cancellationToken)) return false;
        try {
            using var source = new MemoryStream(input, offset, count, writable: false);
            using var deflate = new DeflateStream(source, CompressionMode.Decompress);
            byte[] buffer = new byte[8192];
            int decoded = 0;
            int read;
            while ((read = deflate.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                if (read > expectedCount - decoded) return false;
                decoded += read;
            }
            return decoded == expectedCount;
        } catch (InvalidDataException) {
            return false;
        }
    }

    private static bool TryDecodeStrip(
        byte[] input,
        int inputOffset,
        int inputCount,
        int compression,
        byte[] output,
        int outputOffset,
        int expectedCount,
        CancellationToken cancellationToken) {
        if (input == null || output == null || inputOffset < 0 || inputCount < 0 ||
            outputOffset < 0 || expectedCount < 0 ||
            inputOffset > input.Length - inputCount || outputOffset > output.Length - expectedCount) {
            return false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        switch (compression) {
            case (int)OfficeTiffCompression.None:
                return CopyExact(
                    input, inputOffset, inputCount, output, outputOffset, expectedCount, cancellationToken);
            case (int)OfficeTiffCompression.PackBits:
                return TryDecodePackBits(input, inputOffset, inputCount, output, outputOffset,
                    expectedCount, cancellationToken);
            case (int)OfficeTiffCompression.Lzw:
                return TryDecodeLzw(input, inputOffset, inputCount, output, outputOffset,
                    expectedCount, cancellationToken);
            case (int)OfficeTiffCompression.Deflate:
                return TryDecodeDeflate(input, inputOffset, inputCount, output, outputOffset,
                    expectedCount, allowRawDeflate: false, cancellationToken);
            case 32946:
                return TryDecodeDeflate(input, inputOffset, inputCount, output, outputOffset,
                    expectedCount, allowRawDeflate: true, cancellationToken);
            default:
                return false;
        }
    }

    private static bool TryDecodeDeflate(
        byte[] input,
        int inputOffset,
        int inputCount,
        byte[] output,
        int outputOffset,
        int expectedCount,
        bool allowRawDeflate,
        CancellationToken cancellationToken) {
        var compressed = new byte[inputCount];
        CopyWithCancellation(input, inputOffset, compressed, 0, inputCount, cancellationToken);
        try {
            byte[] inflated = OfficeZlibCodec.Decompress(
                compressed,
                expectedCount,
                expectedCount,
                cancellationToken);
            CopyWithCancellation(inflated, 0, output, outputOffset, expectedCount, cancellationToken);
            return true;
        } catch (Exception exception) when (
            exception is InvalidDataException ||
            exception is OfficeDecompressionSizeLimitException) {
            // Older TIFF writers used raw Deflate under compression tag 32946.
        } catch (NotSupportedException) {
            return false;
        }

        if (!allowRawDeflate) {
            return false;
        }

        if (!OfficeDeflateStreamValidator.TryValidateExact(
                compressed, 0, compressed.Length, expectedCount, cancellationToken)) {
            return false;
        }

        try {
            using var source = new MemoryStream(compressed, writable: false);
            using var deflate = new DeflateStream(source, CompressionMode.Decompress);
            int total = 0;
            while (total < expectedCount) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = deflate.Read(output, outputOffset + total, expectedCount - total);
                if (read == 0) return false;
                total += read;
            }
            return deflate.ReadByte() == -1;
        } catch (InvalidDataException) {
            return false;
        }
    }
}
