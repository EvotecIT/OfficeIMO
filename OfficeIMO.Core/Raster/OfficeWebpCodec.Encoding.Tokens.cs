using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    // Token walks are replayed instead of retaining a pixel-sized token list.
    // A negative distance denotes a cache index (-1 - index); zero is a literal.
    private static void VisitVp8lEncodingTokens(uint[] pixels, int width, Vp8lEncodingMatches matches,
        int cacheBits, CancellationToken cancellationToken, Action<uint, int, int> visit) {
        matches.Reset();
        uint[] cache = cacheBits == 0 ? Array.Empty<uint>() : new uint[1 << cacheBits];
        int position = 0;
        while (position < pixels.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            matches.Find(pixels, position, cancellationToken, out int length, out int distance);
            int count = distance == 0 ? 1 : length;
            if (distance != 0) {
                visit(0, length, GetVp8lEncodingDistance(distance, width));
            } else {
                uint color = pixels[position];
                int cacheIndex = cacheBits == 0 ? -1 : (int)(unchecked(0x1E35A7BDU * color) >> (32 - cacheBits));
                visit(color, 1, cacheIndex >= 0 && cache[cacheIndex] == color ? -1 - cacheIndex : 0);
            }
            // Copies update the cache and hash chain for *every* produced sample,
            // including overlapping copies. This order is part of the VP8L contract.
            for (int index = 0; index < count; index++) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                uint color = pixels[position];
                AddToVp8lCache(cache, cacheBits, color);
                matches.Remember(pixels, position++);
            }
        }
    }

    private static int GetVp8lEncodingDistance(int distance, int width) {
        // Prefer the short two-dimensional plane codes, including row copies.
        for (int code = 1; code <= 120; code++) {
            if (MapVp8lDistance(code, width) == distance) return code;
        }
        return distance + 120;
    }

    /// <summary>Fixed-size hash chains bound both lookback storage and search work.</summary>
    private sealed class Vp8lEncodingMatches {
        private readonly int[] _heads = new int[Vp8lMatchTableMaximumEntries];
        private readonly int[] _previous = new int[Vp8lMatchTableMaximumEntries];

        internal void Reset() {
            Array.Clear(_heads, 0, _heads.Length);
            // Positions in heads are stored plus one. Old ring links are unreachable
            // after heads are cleared and are overwritten before they can be visited.
        }

        internal void Remember(uint[] pixels, int position) {
            if (position + 2 >= pixels.Length) return;
            int hash = Hash(pixels, position);
            _previous[position & (Vp8lMatchTableMaximumEntries - 1)] = _heads[hash];
            _heads[hash] = position + 1;
        }

        internal void Find(uint[] pixels, int position, CancellationToken cancellationToken,
            out int length, out int distance) {
            length = distance = 0;
            if (position + 2 >= pixels.Length) return;
            int candidate = _heads[Hash(pixels, position)] - 1;
            int maximum = Math.Min(4096, pixels.Length - position);
            for (int attempt = 0; attempt < 16 && candidate >= 0; attempt++) {
                int offset = position - candidate;
                if (offset > Vp8lMatchTableMaximumEntries || offset <= 0) break;
                if (pixels[position] == pixels[candidate] &&
                    pixels[position + 1] == pixels[candidate + 1] &&
                    pixels[position + 2] == pixels[candidate + 2]) {
                    int found = 3;
                    while (found < maximum && pixels[position + found] == pixels[candidate + found]) {
                        if ((found & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                        found++;
                    }
                    if (found > length) { length = found; distance = offset; }
                    if (found == maximum) break;
                }
                int next = _previous[candidate & (Vp8lMatchTableMaximumEntries - 1)] - 1;
                if (next >= candidate) break; // never follow an overwritten ring link
                candidate = next;
            }
        }

        private static int Hash(uint[] pixels, int position) {
            uint hash = unchecked(pixels[position] * 0x1E35A7BDU + pixels[position + 1] * 0x9E3779B9U);
            hash = unchecked((hash ^ (hash >> 16)) * 0x85EBCA6BU + pixels[position + 2]);
            return (int)(hash >> 16);
        }
    }

    private static Vp8lEncodingPlan SelectVp8lEncodingPlan(uint[] pixels, int width,
        Vp8lEncodingMatches matches, CancellationToken cancellationToken) {
        Vp8lEncodingPlan plain = CreateVp8lEncodingPlan(pixels, width, matches, 0, cancellationToken);
        Vp8lEncodingPlan cached = CreateVp8lEncodingPlan(pixels, width, matches, 8, cancellationToken);
        return cached.BitCount < plain.BitCount ? cached : plain;
    }

    private static Vp8lEncodingPlan CreateVp8lEncodingPlan(uint[] pixels, int width,
        Vp8lEncodingMatches matches, int cacheBits, CancellationToken cancellationToken) {
        int[][] frequencies = {
            new int[280 + (cacheBits == 0 ? 0 : 1 << cacheBits)],
            new int[256], new int[256], new int[256], new int[40]
        };
        long extraBitCount = 0;
        VisitVp8lEncodingTokens(pixels, width, matches, cacheBits, cancellationToken, (color, length, distance) => {
            if (distance < 0) {
                frequencies[0][279 - distance]++;
            } else if (distance == 0) {
                frequencies[0][(int)(color >> 8) & 255]++;
                frequencies[1][(int)(color >> 16) & 255]++;
                frequencies[2][(int)color & 255]++;
                frequencies[3][(int)(color >> 24)]++;
            } else {
                GetVp8lPrefix(length, 24, out int prefix, out int extra, out _);
                GetVp8lPrefix(distance, 40, out int distancePrefix, out int distanceExtra, out _);
                frequencies[0][256 + prefix]++;
                frequencies[4][distancePrefix]++;
                extraBitCount += extra + distanceExtra;
            }
        });
        var lengths = new byte[5][];
        var counter = new Vp8lBitCounter();
        counter.WriteBits(0, cacheBits == 0 ? 2 : 6); // cache declaration and main meta flag
        for (int channel = 0; channel < lengths.Length; channel++) {
            cancellationToken.ThrowIfCancellationRequested();
            lengths[channel] = OfficeHuffmanCodeLengths.Create(frequencies[channel], 15);
            WriteVp8lHuffmanTree(counter, lengths[channel]);
            int used = 0;
            foreach (byte value in lengths[channel]) if (value > 0) used++;
            if (used > 1) {
                for (int symbol = 0; symbol < lengths[channel].Length; symbol++) {
                    extraBitCount += (long)frequencies[channel][symbol] * lengths[channel][symbol];
                }
            }
        }
        return new Vp8lEncodingPlan(cacheBits, lengths, counter.BitCount + extraBitCount);
    }

    private static void WriteVp8lEncodingImage(ILsbBitWriter writer, uint[] pixels, int width,
        Vp8lEncodingMatches matches, Vp8lEncodingPlan plan, bool mainImage, CancellationToken cancellationToken) {
        writer.WriteBits(plan.CacheBits == 0 ? 0U : 1U, 1);
        if (plan.CacheBits != 0) writer.WriteBits((uint)plan.CacheBits, 4);
        if (mainImage) writer.WriteBits(0, 1); // one statistical prefix group
        var codes = new Vp8lCodebook[5];
        for (int channel = 0; channel < codes.Length; channel++) {
            cancellationToken.ThrowIfCancellationRequested();
            WriteVp8lHuffmanTree(writer, plan.Lengths[channel]);
            codes[channel] = new Vp8lCodebook(plan.Lengths[channel]);
        }
        VisitVp8lEncodingTokens(pixels, width, matches, plan.CacheBits, cancellationToken, (color, length, distance) => {
            if (distance < 0) {
                codes[0].Write(writer, 279 - distance);
            } else if (distance == 0) {
                codes[0].Write(writer, (int)(color >> 8) & 255);
                codes[1].Write(writer, (int)(color >> 16) & 255);
                codes[2].Write(writer, (int)color & 255);
                codes[3].Write(writer, (int)(color >> 24));
            } else {
                GetVp8lPrefix(length, 24, out int prefix, out int extra, out int value);
                codes[0].Write(writer, 256 + prefix);
                writer.WriteBits((uint)value, extra);
                GetVp8lPrefix(distance, 40, out int distancePrefix, out int distanceExtra, out int distanceValue);
                codes[4].Write(writer, distancePrefix);
                writer.WriteBits((uint)distanceValue, distanceExtra);
            }
        });
    }

    private sealed class Vp8lEncodingPlan {
        internal Vp8lEncodingPlan(int cacheBits, byte[][] lengths, long bitCount) {
            CacheBits = cacheBits;
            Lengths = lengths;
            BitCount = bitCount;
        }
        internal int CacheBits { get; }
        internal byte[][] Lengths { get; }
        internal long BitCount { get; }
    }

    private sealed class Vp8lBitCounter : ILsbBitWriter {
        internal long BitCount { get; private set; }
        public void WriteBits(uint value, int count) => BitCount += count;
        public void Flush() { }
    }
}
