using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private const int Vp8lGeneralMaximumPixels = 16_000_000;
    private static readonly sbyte[] Vp8lDistanceMap = {
         0,1,  1,0,  1,1, -1,1,  0,2,  2,0,  1,2, -1,2,  2,1, -2,1,
         2,2, -2,2,  0,3,  3,0,  1,3, -1,3,  3,1, -3,1,  2,3, -2,3,
         3,2, -3,2,  0,4,  4,0,  1,4, -1,4,  4,1, -4,1,  3,3, -3,3,
         2,4, -2,4,  4,2, -4,2,  0,5,  3,4, -3,4,  4,3, -4,3,  5,0,
         1,5, -1,5,  5,1, -5,1,  2,5, -2,5,  5,2, -5,2,  4,4, -4,4,
         3,5, -3,5,  5,3, -5,3,  0,6,  6,0,  1,6, -1,6,  6,1, -6,1,
         2,6, -2,6,  6,2, -6,2,  4,5, -4,5,  5,4, -5,4,  3,6, -3,6,
         6,3, -6,3,  0,7,  7,0,  1,7, -1,7,  5,5, -5,5,  7,1, -7,1,
         4,6, -4,6,  6,4, -6,4,  2,7, -2,7,  7,2, -7,2,  3,7, -3,7,
         7,3, -7,3,  5,6, -5,6,  6,5, -6,5,  8,0,  4,7, -4,7,  7,4,
        -7,4,  8,1,  8,2,  6,6, -6,6,  8,3,  5,7, -5,7,  7,5, -7,5,
         8,4,  6,7, -6,7,  7,6, -7,6,  8,5,  7,7, -7,7,  8,6,  8,7
    };

    private static bool TryDecodeVp8l(
        byte[]? encodedBytes,
        CancellationToken cancellationToken,
        out OfficeRasterImage? image) {
        image = null;
        cancellationToken.ThrowIfCancellationRequested();
        if (!IsWebp(encodedBytes) || encodedBytes == null || encodedBytes.Length < 22 ||
            encodedBytes.Length > OfficeRasterGuards.MaximumEncodedBytes) return false;
        try {
            if (ReadUInt32(encodedBytes, 4) != encodedBytes.Length - 8 ||
                !TryFindChunk(encodedBytes, "VP8L", out int payloadOffset, out int payloadLength) ||
                payloadLength < 5 || encodedBytes[payloadOffset] != 0x2F) return false;
            var reader = new LsbBitReader(encodedBytes, payloadOffset + 1, payloadLength - 1);
            int width = checked((int)reader.ReadBits(14) + 1);
            int height = checked((int)reader.ReadBits(14) + 1);
            reader.ReadBits(1);
            if (reader.ReadBits(3) != 0 ||
                !OfficeRasterGuards.TryEnsurePixelCount(width, height, out int pixels) ||
                pixels > Vp8lGeneralMaximumPixels) return false;

            var allocationBudget = new Vp8lAllocationBudget();
            if (!allocationBudget.TryReserveBytes(encodedBytes.Length) ||
                !TryReadVp8lTransforms(reader, width, height, allocationBudget, cancellationToken, out int encodedWidth,
                    out List<Vp8lTransform> transforms)) return false;
            if (!TryDecodeVp8lImageData(reader, encodedWidth, height, allowMetaCodes: true, 0,
                    allocationBudget, cancellationToken, out uint[] packed))
                return false;
            if (!TryApplyVp8lTransforms(packed, encodedWidth, height, width, transforms,
                    allocationBudget, cancellationToken, out uint[] argb))
                return false;
            if (argb.Length != pixels || !reader.HasOnlyZeroPadding()) return false;

            if (!allocationBudget.TryReserveArray(pixels, sizeof(uint))) return false;
            byte[] rgba = OfficeRasterGuards.AllocateRgba32(width, height, "WebP decoded pixels exceed the managed limit.");
            for (int pixel = 0; pixel < argb.Length; pixel++) {
                if ((pixel & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                uint color = argb[pixel];
                int offset = pixel * 4;
                rgba[offset] = (byte)(color >> 16);
                rgba[offset + 1] = (byte)(color >> 8);
                rgba[offset + 2] = (byte)color;
                rgba[offset + 3] = (byte)(color >> 24);
            }
            image = OfficeRasterImage.FromOwnedRgba32(width, height, rgba);
            return true;
        } catch (Exception exception) when (
            exception is ArgumentException || exception is FormatException || exception is OverflowException ||
            exception is OutOfMemoryException) {
            image = null;
            return false;
        }
    }

    private static bool TryDecodeVp8lImageData(
        LsbBitReader reader,
        int width,
        int height,
        bool allowMetaCodes,
        int depth,
        Vp8lAllocationBudget allocationBudget,
        CancellationToken cancellationToken,
        out uint[] pixels) {
        pixels = Array.Empty<uint>();
        if (depth > 4 || !OfficeRasterGuards.TryEnsurePixelCount(width, height, out int pixelCount) ||
            pixelCount > Vp8lGeneralMaximumPixels) return false;
        int cacheBits = 0;
        if (reader.ReadBits(1) != 0) {
            cacheBits = (int)reader.ReadBits(4);
            if (cacheBits < 1 || cacheBits > 11) return false;
        }
        int cacheSize = cacheBits == 0 ? 0 : 1 << cacheBits;

        int prefixBits = 0;
        int prefixWidth = 0;
        uint[]? prefixImage = null;
        int groupCount = 1;
        if (allowMetaCodes && reader.ReadBits(1) != 0) {
            prefixBits = (int)reader.ReadBits(3) + 2;
            prefixWidth = DivideRoundUp(width, 1 << prefixBits);
            int prefixHeight = DivideRoundUp(height, 1 << prefixBits);
            if (!TryDecodeVp8lImageData(reader, prefixWidth, prefixHeight, false, depth + 1,
                    allocationBudget, cancellationToken, out prefixImage))
                return false;
            int maximumGroup = 0;
            for (int index = 0; index < prefixImage.Length; index++) {
                int group = (int)((prefixImage[index] >> 8) & 0xFFFFU);
                if (group > maximumGroup) maximumGroup = group;
            }
            groupCount = checked(maximumGroup + 1);
            if (groupCount > 65536) return false;
        }

        if (!allocationBudget.TryReserveArray(groupCount, IntPtr.Size) ||
            !allocationBudget.TryReserveBytes((long)groupCount * 64L)) return false;
        var groups = new Vp8lHuffmanGroup[groupCount];
        for (int group = 0; group < groupCount; group++) {
            if ((group & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (!TryReadHuffmanTree(reader, 280 + cacheSize, allocationBudget, out Vp8lHuffmanTree green))
                return false;
            if (!TryReadHuffmanTree(reader, 256, allocationBudget, out Vp8lHuffmanTree red))
                return false;
            if (!TryReadHuffmanTree(reader, 256, allocationBudget, out Vp8lHuffmanTree blue))
                return false;
            if (!TryReadHuffmanTree(reader, 256, allocationBudget, out Vp8lHuffmanTree alpha))
                return false;
            if (!TryReadHuffmanTree(reader, 40, allocationBudget, out Vp8lHuffmanTree distance))
                return false;
            groups[group] = new Vp8lHuffmanGroup(green, red, blue, alpha, distance);
        }

        if (!allocationBudget.TryReserveArray(pixelCount, sizeof(uint)) ||
            cacheSize != 0 && !allocationBudget.TryReserveArray(cacheSize, sizeof(uint))) return false;
        pixels = new uint[pixelCount];
        uint[] cache = cacheSize == 0 ? Array.Empty<uint>() : new uint[cacheSize];
        int position = 0;
        while (position < pixelCount) {
            if ((position & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int x = position % width;
            int y = position / width;
            int groupIndex = prefixImage == null ? 0 :
                (int)((prefixImage[(y >> prefixBits) * prefixWidth + (x >> prefixBits)] >> 8) & 0xFFFFU);
            if ((uint)groupIndex >= (uint)groups.Length) return false;
            Vp8lHuffmanGroup group = groups[groupIndex];
            int symbol = group.Green.ReadSymbol(reader);
            if (symbol < 0) return false;
            if (symbol < 256) {
                int red = group.Red.ReadSymbol(reader);
                int blue = group.Blue.ReadSymbol(reader);
                int alpha = group.Alpha.ReadSymbol(reader);
                if ((uint)red > 255U || (uint)blue > 255U || (uint)alpha > 255U)
                    return false;
                uint color = (uint)(alpha << 24 | red << 16 | symbol << 8 | blue);
                pixels[position++] = color;
                AddToVp8lCache(cache, cacheBits, color);
            } else if (symbol < 280) {
                int length = ReadVp8lPrefixValue(reader, symbol - 256);
                int distancePrefix = group.Distance.ReadSymbol(reader);
                if (length < 1 || distancePrefix < 0 || distancePrefix >= 40)
                    return false;
                int distanceCode = ReadVp8lPrefixValue(reader, distancePrefix);
                int distance = MapVp8lDistance(distanceCode, width);
                if (distance < 1 || distance > position || length > pixelCount - position)
                    return false;
                for (int index = 0; index < length; index++) {
                    uint color = pixels[position - distance];
                    pixels[position++] = color;
                    AddToVp8lCache(cache, cacheBits, color);
                }
            } else {
                int cacheIndex = symbol - 280;
                if ((uint)cacheIndex >= (uint)cache.Length) return false;
                uint color = cache[cacheIndex];
                pixels[position++] = color;
                AddToVp8lCache(cache, cacheBits, color);
            }
        }
        return true;
    }

    private static int ReadVp8lPrefixValue(LsbBitReader reader, int prefix) {
        if (prefix < 4) return prefix + 1;
        int extraBits = (prefix - 2) >> 1;
        int offset = (2 + (prefix & 1)) << extraBits;
        return checked(offset + (int)reader.ReadBits(extraBits) + 1);
    }

    private static int MapVp8lDistance(int distanceCode, int width) {
        if (distanceCode > 120) return distanceCode - 120;
        if (distanceCode < 1) return 0;
        int index = (distanceCode - 1) * 2;
        int distance = Vp8lDistanceMap[index] + Vp8lDistanceMap[index + 1] * width;
        return Math.Max(1, distance);
    }

    private static void AddToVp8lCache(uint[] cache, int cacheBits, uint color) {
        if (cacheBits == 0) return;
        int index = (int)((0x1E35A7BDU * color) >> (32 - cacheBits));
        cache[index] = color;
    }

    private static int DivideRoundUp(int value, int divisor) => checked((value + divisor - 1) / divisor);

    internal sealed class Vp8lAllocationBudget {
        private long _reservedBytes;

        internal bool TryReserveArray(long elements, int elementSize) {
            if (elements < 0 || elementSize < 1) return false;
            return TryReserveBytes(checked(elements * elementSize + 24L));
        }

        internal bool TryReserveBytes(long bytes) {
            if (bytes < 0 || _reservedBytes > OfficeRasterGuards.MaximumDecodedBytes - bytes) return false;
            _reservedBytes += bytes;
            return true;
        }
    }

    private sealed class Vp8lHuffmanGroup {
        internal Vp8lHuffmanGroup(Vp8lHuffmanTree green, Vp8lHuffmanTree red,
            Vp8lHuffmanTree blue, Vp8lHuffmanTree alpha, Vp8lHuffmanTree distance) {
            Green = green;
            Red = red;
            Blue = blue;
            Alpha = alpha;
            Distance = distance;
        }
        internal Vp8lHuffmanTree Green { get; }
        internal Vp8lHuffmanTree Red { get; }
        internal Vp8lHuffmanTree Blue { get; }
        internal Vp8lHuffmanTree Alpha { get; }
        internal Vp8lHuffmanTree Distance { get; }
    }
}
