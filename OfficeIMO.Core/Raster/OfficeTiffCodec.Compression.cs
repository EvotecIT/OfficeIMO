using System;
using System.IO;
using System.IO.Compression;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private static bool TryDecodeStrip(
        byte[] input,
        int inputOffset,
        int inputCount,
        int compression,
        byte[] output,
        int outputOffset,
        int expectedCount) {
        switch (compression) {
            case (int)OfficeTiffCompression.None:
                return CopyExact(input, inputOffset, inputCount, output, outputOffset, expectedCount);
            case (int)OfficeTiffCompression.PackBits:
                return TryDecodePackBits(input, inputOffset, inputCount, output, outputOffset, expectedCount);
            case (int)OfficeTiffCompression.Deflate:
            case 32946:
                return TryDecodeDeflate(input, inputOffset, inputCount, output, outputOffset, expectedCount);
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
        int expectedCount) {
        var compressed = new byte[inputCount];
        Buffer.BlockCopy(input, inputOffset, compressed, 0, inputCount);
        try {
            byte[] inflated = OfficeZlibCodec.Decompress(
                compressed,
                expectedCount,
                expectedCount);
            Buffer.BlockCopy(inflated, 0, output, outputOffset, expectedCount);
            return true;
        } catch (InvalidDataException) {
            // Older TIFF writers used raw Deflate under compression tag 32946.
        } catch (NotSupportedException) {
            return false;
        }

        try {
            using var source = new MemoryStream(compressed, writable: false);
            using var deflate = new DeflateStream(source, CompressionMode.Decompress);
            int total = 0;
            while (total < expectedCount) {
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
