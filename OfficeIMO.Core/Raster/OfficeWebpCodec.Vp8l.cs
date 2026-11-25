#if NET8_0_OR_GREATER
using Vp8lPixels = System.Span<uint>;
#else
using Vp8lPixels = uint[];
#endif

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
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
        long retainedManagedBytes,
        out OfficeRasterImage? image) {
        image = null;
        cancellationToken.ThrowIfCancellationRequested();
        if (!IsWebp(encodedBytes) || encodedBytes == null || encodedBytes.Length < 22 ||
            encodedBytes.Length > OfficeRasterGuards.MaximumEncodedBytes) return false;
        try {
            if (ReadUInt32(encodedBytes, 4) != encodedBytes.Length - 8 ||
                !TryFindChunk(encodedBytes, "VP8L", cancellationToken, out int payloadOffset, out int payloadLength) ||
                payloadLength < 5 || encodedBytes[payloadOffset] != 0x2F) return false;
            var reader = new LsbBitReader(encodedBytes, payloadOffset + 1, payloadLength - 1);
            int width = checked((int)reader.ReadBits(14) + 1);
            int height = checked((int)reader.ReadBits(14) + 1);
            reader.ReadBits(1);
            if (reader.ReadBits(3) != 0 ||
                !OfficeRasterGuards.TryEnsurePixelCount(width, height, out int pixels)) return false;

            var allocationBudget = new Vp8lAllocationBudget();
            if (!allocationBudget.TryReserveBytes(retainedManagedBytes) ||
                !allocationBudget.TryReserveArray(encodedBytes.Length, sizeof(byte))) return false;

            if (!TryDecodeVp8lArgb(reader, width, height, allocationBudget, cancellationToken, writeRgba: true,
                    out _, out byte[]? rgba) || rgba == null) return false;
            image = OfficeRasterImage.FromOwnedRgba32(width, height, rgba);
            return true;
        } catch (Exception exception) when (
            exception is ArgumentException || exception is FormatException || exception is OverflowException ||
            exception is OutOfMemoryException) {
            image = null;
            return false;
        }
    }

    /// <summary>
    /// Decodes a headerless lossless image stream used by VP8L and compressed ALPH chunks. With
    /// <paramref name="writeRgba"/>, the entropy-coded pixels decode into the tail of the RGBA output
    /// itself and the transforms restore them in place, so the image needs one full-size buffer.
    /// </summary>
    private static bool TryDecodeVp8lArgb(LsbBitReader reader, int width, int height,
        Vp8lAllocationBudget allocationBudget, CancellationToken cancellationToken, bool writeRgba,
        out uint[] argb, out byte[]? rgba) {
        argb = Array.Empty<uint>();
        rgba = null;
        if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out int pixels) ||
            !TryReadVp8lTransforms(reader, width, height, allocationBudget, cancellationToken, out int encodedWidth,
                out List<Vp8lTransform> transforms) ||
            !TryReadVp8lImageHeader(reader, encodedWidth, height, allowMetaCodes: true, 0,
                allocationBudget, cancellationToken, out Vp8lImageHeader header)) return false;
        uint[]? packedArray = null;
        Vp8lPixels packed;
#if NET8_0_OR_GREATER
        if (writeRgba) {
            // Transforms never widen the stream past the image, so the packed pixels fit the output.
            if (!allocationBudget.TryReserveArray(pixels, sizeof(uint))) return false;
            cancellationToken.ThrowIfCancellationRequested();
            rgba = OfficeRasterGuards.AllocateRgba32Uninitialized(width, height, "WebP decoded pixels exceed the managed limit.");
            int packedLength = checked(encodedWidth * height);
            packed = System.Runtime.InteropServices.MemoryMarshal.Cast<byte, uint>(rgba.AsSpan()).Slice(pixels - packedLength, packedLength);
        } else
#endif
        {
            if (!header.TryAllocatePixels(allocationBudget, cancellationToken, out packedArray)) return false;
            packed = packedArray;
        }
#if !NET8_0_OR_GREATER
        if (writeRgba) {
            // Without spans the packed pixels keep their own buffer beside the output.
            if (!allocationBudget.TryReserveArray(pixels, sizeof(uint))) return false;
            cancellationToken.ThrowIfCancellationRequested();
            rgba = OfficeRasterGuards.AllocateRgba32Uninitialized(width, height, "WebP decoded pixels exceed the managed limit.");
        }
#endif
        if (!header.TryDecodePixels(reader, packed, cancellationToken) ||
            // The transforms read nothing more from the stream, so reject trailing data before restoring pixels.
            !reader.HasOnlyZeroPadding(cancellationToken)) {
            rgba = null;
            return false;
        }
        if (!TryApplyVp8lTransforms(packed, packedArray, encodedWidth, height, width, transforms,
                allocationBudget, cancellationToken, rgba, out argb)) {
            rgba = null;
            return false;
        }
        return rgba != null || argb.Length == pixels;
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
        if (!TryReadVp8lImageHeader(reader, width, height, allowMetaCodes, depth, allocationBudget, cancellationToken,
                out Vp8lImageHeader header) ||
            !header.TryAllocatePixels(allocationBudget, cancellationToken, out uint[] decoded) ||
            !header.TryDecodePixels(reader, decoded, cancellationToken)) return false;
        pixels = decoded;
        return true;
    }

    /// <summary>Reads the color cache, meta prefix codes and Huffman groups that precede an image's pixels.</summary>
    private static bool TryReadVp8lImageHeader(
        LsbBitReader reader,
        int width,
        int height,
        bool allowMetaCodes,
        int depth,
        Vp8lAllocationBudget allocationBudget,
        CancellationToken cancellationToken,
        out Vp8lImageHeader header) {
        header = default;
        if (depth > 4 || !OfficeRasterGuards.TryEnsurePixelCount(width, height, out int pixelCount)) return false;
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
            int maximumGroup = FindMaximumVp8lGroup(prefixImage, cancellationToken);
            groupCount = checked(maximumGroup + 1);
            if (groupCount > 65536) return false;
        }

        if (!allocationBudget.TryReserveArray(groupCount, IntPtr.Size) ||
            !allocationBudget.TryReserveBytes((long)groupCount * 64L)) return false;
        cancellationToken.ThrowIfCancellationRequested();
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

        if (cacheSize != 0 && !allocationBudget.TryReserveArray(cacheSize, sizeof(uint))) return false;
        header = new Vp8lImageHeader(width, pixelCount, groups, prefixImage, prefixBits, prefixWidth, cacheBits);
        return true;
    }

    /// <summary>The coding parameters of one entropy-coded image, read before its pixels.</summary>
    private readonly struct Vp8lImageHeader {
        private readonly int _width;
        private readonly int _pixelCount;
        private readonly Vp8lHuffmanGroup[] _groups;
        private readonly uint[]? _prefixImage;
        private readonly int _prefixBits;
        private readonly int _prefixWidth;
        private readonly int _cacheBits;

        internal Vp8lImageHeader(int width, int pixelCount, Vp8lHuffmanGroup[] groups, uint[]? prefixImage,
            int prefixBits, int prefixWidth, int cacheBits) {
            _width = width;
            _pixelCount = pixelCount;
            _groups = groups;
            _prefixImage = prefixImage;
            _prefixBits = prefixBits;
            _prefixWidth = prefixWidth;
            _cacheBits = cacheBits;
        }

        internal bool TryAllocatePixels(Vp8lAllocationBudget allocationBudget, CancellationToken cancellationToken,
            out uint[] pixels) {
            pixels = Array.Empty<uint>();
            if (!allocationBudget.TryReserveArray(_pixelCount, sizeof(uint))) return false;
            cancellationToken.ThrowIfCancellationRequested();
            // Every position is written before the loop ends or the decode fails.
            pixels = AllocateVp8lPixels(_pixelCount);
            return true;
        }

        internal bool TryDecodePixels(LsbBitReader reader, Vp8lPixels pixels, CancellationToken cancellationToken) {
            if (pixels.Length != _pixelCount) return false;
            cancellationToken.ThrowIfCancellationRequested();
            uint[] cache = _cacheBits == 0 ? Array.Empty<uint>() : new uint[1 << _cacheBits];
            return TryDecodeVp8lPixels(reader, pixels, _width, _groups, _prefixImage, _prefixBits, _prefixWidth,
                cache, _cacheBits, cancellationToken);
        }
    }

    /// <summary>
    /// Decodes the entropy-coded pixels. The loop keeps only the bit state, the active group, the
    /// output and the position in registers; group changes, cancellation, cache lookups and
    /// backreference distances go through <see cref="Vp8lPixelState"/> on their less frequent paths.
    /// </summary>
    private static bool TryDecodeVp8lPixels(LsbBitReader reader, Vp8lPixels pixels, int width,
        Vp8lHuffmanGroup[] groups, uint[]? prefixImage, int prefixBits, int prefixWidth,
        uint[] cache, int cacheBits, CancellationToken cancellationToken) {
        var state = new Vp8lPixelState(width, groups, prefixImage, prefixBits, prefixWidth, cache, cacheBits, cancellationToken);
        Vp8lHuffmanGroup group = groups[0];
        int[] greenTable = group.Green.Table;
        int position = 0;
        int boundary = 0;
        byte[] input = reader.Input;
        int end = reader.End;
        int offset = reader.Offset;
        ulong buffer = reader.Buffer;
        int count = reader.BitCount;
        // Refills and lookups take and return values, so the bit state stays in registers.
        while (position < pixels.Length) {
            if (position >= boundary) {
                if (count < 0 || !state.TryAdvance(position, ref group, out boundary)) return false;
                greenTable = group.Green.Table;
            }
            if ((uint)count < 32U && offset < end) {
                buffer |= LsbBitReader.LoadWord(input, offset, end) << count;
                int bytes = Math.Min((63 - count) >> 3, end - offset);
                offset += bytes;
                count += bytes << 3;
            }
            int entry = Vp8lHuffmanTree.Lookup(greenTable, buffer);
            buffer >>= entry & 0xFF;
            count -= entry & 0xFF;
            int symbol = entry >> 8;
            if (symbol < 256) {
                if (group.HasConstantRedBlueAlpha) {
                    pixels[position++] = group.ConstantRedBlueAlpha | (uint)symbol << 8;
                    continue;
                }
                // Codes are at most 15 bits, so 32 buffered bits cover two of them.
                if ((uint)count < 32U && offset < end) {
                    buffer |= LsbBitReader.LoadWord(input, offset, end) << count;
                    int bytes = Math.Min((63 - count) >> 3, end - offset);
                    offset += bytes;
                    count += bytes << 3;
                }
                entry = Vp8lHuffmanTree.Lookup(group.Red.Table, buffer);
                buffer >>= entry & 0xFF;
                count -= entry & 0xFF;
                uint color = (uint)(entry >> 8) << 16 | (uint)symbol << 8;
                entry = Vp8lHuffmanTree.Lookup(group.Blue.Table, buffer);
                buffer >>= entry & 0xFF;
                count -= entry & 0xFF;
                color |= (uint)(entry >> 8);
                if ((uint)count < 32U && offset < end) {
                    buffer |= LsbBitReader.LoadWord(input, offset, end) << count;
                    int bytes = Math.Min((63 - count) >> 3, end - offset);
                    offset += bytes;
                    count += bytes << 3;
                }
                entry = Vp8lHuffmanTree.Lookup(group.Alpha.Table, buffer);
                buffer >>= entry & 0xFF;
                count -= entry & 0xFF;
                // Red, blue and alpha alphabets have 256 symbols, so each value fits its byte.
                pixels[position++] = color | (uint)(entry >> 8) << 24;
            } else if (symbol < 280) {
                reader.Resume(offset, buffer, count);
                int length = ReadVp8lPrefixValue(reader, symbol - 256);
                int distancePrefix = group.Distance.ReadSymbol(reader);
                int distanceCode = ReadVp8lPrefixValue(reader, distancePrefix);
                offset = reader.Offset;
                buffer = reader.Buffer;
                count = reader.BitCount;
                if (distancePrefix >= 40 || !state.TryCopy(pixels, position, distanceCode, length)) return false;
                position += length;
            } else {
                if (!state.TryReadCache(pixels, position, symbol - 280, out uint color)) return false;
                pixels[position++] = color;
            }
        }
        if (count < 0) return false;
        reader.Resume(offset, buffer, count);
        return true;
    }

    /// <summary>State the pixel loop needs only at group boundaries, cache codes and backreferences.</summary>
    private sealed class Vp8lPixelState {
        private readonly int _width;
        private readonly Vp8lHuffmanGroup[] _groups;
        private readonly uint[]? _prefixImage;
        private readonly int _prefixBits;
        private readonly int _prefixWidth;
        private readonly uint[] _cache;
        private readonly int _cacheShift;
        private readonly CancellationToken _cancellationToken;
        private int _nextCancellationCheck;
        private int _nextGroupPosition;
        // Cache inserts are deferred until a cache code needs them; inserting the same colors in
        // the same order just before each lookup leaves the cache in the state the format defines.
        private int _cached;

        internal Vp8lPixelState(int width, Vp8lHuffmanGroup[] groups, uint[]? prefixImage, int prefixBits,
            int prefixWidth, uint[] cache, int cacheBits, CancellationToken cancellationToken) {
            _width = width;
            _groups = groups;
            _prefixImage = prefixImage;
            _prefixBits = prefixBits;
            _prefixWidth = prefixWidth;
            _cache = cache;
            _cacheShift = 32 - cacheBits;
            _cancellationToken = cancellationToken;
            _nextGroupPosition = prefixImage == null ? int.MaxValue : 0;
        }

        /// <summary>
        /// Checks cancellation and selects the group at <paramref name="position"/> when either is
        /// due; returns the next position at which either is due again.
        /// </summary>
        internal bool TryAdvance(int position, ref Vp8lHuffmanGroup group, out int boundary) {
            // Backreferences advance by several pixels and can skip exact multiples.
            if (position >= _nextCancellationCheck) {
                _cancellationToken.ThrowIfCancellationRequested();
                _nextCancellationCheck = position + 4096;
            }
            if (position >= _nextGroupPosition) {
                int y = position / _width;
                int x = position - y * _width;
                int groupIndex = (int)((_prefixImage![(y >> _prefixBits) * _prefixWidth + (x >> _prefixBits)] >> 8) & 0xFFFFU);
                if ((uint)groupIndex >= (uint)_groups.Length) {
                    boundary = 0;
                    return false;
                }
                group = _groups[groupIndex];
                // Reuse this group up to the next block or row boundary.
                // A copy may jump past it; its next encoded symbol selects
                // the group at the new position, not inside copied pixels.
                int blockSize = 1 << _prefixBits;
                _nextGroupPosition = position + Math.Min(_width - x, blockSize - (x & (blockSize - 1)));
            }
            boundary = Math.Min(_nextCancellationCheck, _nextGroupPosition);
            return true;
        }

        internal bool TryCopy(Vp8lPixels pixels, int position, int distanceCode, int length) {
            int distance = MapVp8lDistance(distanceCode, _width);
            if (distance < 1 || distance > position || length > pixels.Length - position) return false;
            CopyVp8lBackReference(pixels, position, distance, length);
            return true;
        }

        internal bool TryReadCache(Vp8lPixels pixels, int position, int index, out uint color) {
            uint[] cache = _cache;
            if ((uint)index >= (uint)cache.Length) {
                color = 0;
                return false;
            }
            int shift = _cacheShift;
            int cached = _cached;
            for (; cached < position; cached++) {
                uint value = pixels[cached];
                cache[(int)((0x1E35A7BDU * value) >> shift)] = value;
            }
            _cached = cached;
            color = cache[index];
            return true;
        }
    }

    /// <summary>Copies an LZ77 backreference whose source may overlap the destination.</summary>
    private static void CopyVp8lBackReference(Vp8lPixels pixels, int position, int distance, int length) {
        int source = position - distance;
        if (length <= 16) {
            // A forward pixel loop repeats overlapping patterns and avoids call overhead on short runs.
            for (int index = 0; index < length; index++) pixels[position + index] = pixels[source + index];
            return;
        }
        if (distance >= length) {
            CopyVp8lPixels(pixels, source, pixels, position, length);
            return;
        }
        if (distance == 1) {
            uint color = pixels[source];
#if NET8_0_OR_GREATER
            pixels.Slice(position, length).Fill(color);
#else
            for (int index = 0; index < length; index++) pixels[position + index] = color;
#endif
            return;
        }
        // The copied span repeats with the backreference period. Copying from its start while the
        // written length stays a whole number of periods doubles the span on each pass.
        int copied = 0;
        int available = distance;
        while (copied < length) {
            int chunk = Math.Min(available, length - copied);
            CopyVp8lPixels(pixels, source, pixels, position + copied, chunk);
            copied += chunk;
            available += chunk;
        }
    }

    private static void CopyVp8lPixels(Vp8lPixels source, int sourceIndex, Vp8lPixels destination, int destinationIndex, int length) {
#if NET8_0_OR_GREATER
        source.Slice(sourceIndex, length).CopyTo(destination.Slice(destinationIndex, length));
#else
        Array.Copy(source, sourceIndex, destination, destinationIndex, length);
#endif
    }

    internal static int FindMaximumVp8lGroup(
        uint[] prefixImage,
        CancellationToken cancellationToken) {
        int maximumGroup = 0;
        for (int index = 0; index < prefixImage.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int group = (int)((prefixImage[index] >> 8) & 0xFFFFU);
            if (group > maximumGroup) maximumGroup = group;
        }
        return maximumGroup;
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

        // Green alphabets reach 280 + 2048 symbols with the largest color cache.
        private const int MaximumAlphabet = 280 + 2048;
        private int[]?[]? _singleSymbolTables;

        /// <summary>
        /// Returns the prefix table for a code with one symbol: a root level whose entries all decode that
        /// symbol with zero bits. Images can declare up to 65,536 groups of five codes, so one table per
        /// symbol is shared for the whole decode and reserved once.
        /// </summary>
        internal bool TryGetSingleSymbolTable(int symbol, out int[] table) {
            table = Array.Empty<int>();
            if ((uint)symbol >= MaximumAlphabet) return false;
            if (_singleSymbolTables == null) {
                if (!TryReserveArray(MaximumAlphabet, IntPtr.Size)) return false;
                _singleSymbolTables = new int[]?[MaximumAlphabet];
            }
            int[]? shared = _singleSymbolTables[symbol];
            if (shared == null) {
                if (!TryReserveArray(1 << Vp8lHuffmanTree.RootBits, sizeof(int))) return false;
                shared = new int[1 << Vp8lHuffmanTree.RootBits];
#if NET8_0_OR_GREATER
                Array.Fill(shared, symbol << 8);
#else
                for (int index = 0; index < shared.Length; index++) shared[index] = symbol << 8;
#endif
                _singleSymbolTables[symbol] = shared;
            }
            table = shared;
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
            if (red.IsSingleSymbol && blue.IsSingleSymbol && alpha.IsSingleSymbol) {
                HasConstantRedBlueAlpha = true;
                ConstantRedBlueAlpha = (uint)((alpha.Table[0] >> 8) << 24 | (red.Table[0] >> 8) << 16 | (blue.Table[0] >> 8));
            }
        }
        /// <summary>Gets whether literals take red, blue and alpha from single-symbol codes that read no bits.</summary>
        internal bool HasConstantRedBlueAlpha { get; }
        /// <summary>Gets the red, blue and alpha bits shared by every literal when <see cref="HasConstantRedBlueAlpha"/> is set.</summary>
        internal uint ConstantRedBlueAlpha { get; }
        internal Vp8lHuffmanTree Green { get; }
        internal Vp8lHuffmanTree Red { get; }
        internal Vp8lHuffmanTree Blue { get; }
        internal Vp8lHuffmanTree Alpha { get; }
        internal Vp8lHuffmanTree Distance { get; }
    }
}
