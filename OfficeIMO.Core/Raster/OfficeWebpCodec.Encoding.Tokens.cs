namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    // Distance codes end at prefix 39 (1,048,576), so a copy reaches at most this far back.
    private const int Vp8lMaximumCopyDistance = (1 << 20) - 120;
    private const int Vp8lMaximumCopyLength = 4096;
    private const int Vp8lMaximumMatchHashBits = 18;

    /// <summary>
    /// The LZ77 parse of one image: literals and copies in stream order. Cache indices are not stored;
    /// they depend on the cache size, so every pass that needs them replays the pixels.
    /// </summary>
    private sealed class Vp8lTokens {
        private const ulong CopyFlag = 1UL << 63;
        private ulong[] _items;

        internal Vp8lTokens(int capacity) => _items = new ulong[Math.Max(16, capacity)];

        // Takes ownership of tokens already in stream order.
        internal Vp8lTokens(ulong[] items, int count) {
            _items = items;
            Count = count;
        }

        internal static ulong Literal(uint color) => color;

        internal static ulong Copy(int length, int distanceCode) => CopyFlag | (ulong)length << 32 | (uint)distanceCode;

        internal int Count { get; private set; }

        internal void AddLiteral(uint color) => Add(Literal(color));

        internal void AddCopy(int length, int distanceCode) => Add(Copy(length, distanceCode));

        internal static bool IsCopy(ulong token) => (token & CopyFlag) != 0;

        internal static uint Color(ulong token) => (uint)token;

        internal static int Length(ulong token) => IsCopy(token) ? (int)(token >> 32) & 0x1FFF : 1;

        internal static int DistanceCode(ulong token) => (int)(uint)token;

        internal ulong this[int index] => _items[index];

        private void Add(ulong token) {
            if (Count == _items.Length) {
                var grown = new ulong[checked(_items.Length * 2)];
                Array.Copy(_items, grown, Count);
                _items = grown;
            }
            _items[Count++] = token;
        }
    }

    /// <summary>Hash chains over the whole copy window, plus the two distances that cost least to code.</summary>
    private sealed class Vp8lMatcher {
        private int[] _heads;
        private readonly int _hashShift;
        private readonly uint[] _pixels;
        private int[] _previous;
        private readonly int _width;
        private readonly int _attempts;
        private readonly int[] _planeCodes;

        internal Vp8lMatcher(uint[] pixels, int width, int attempts) {
            int hashBits = GetVp8lMatchHashBits(pixels.Length);
            _heads = new int[1 << hashBits];
            _hashShift = 32 - hashBits;
            _pixels = pixels;
            _previous = new int[pixels.Length];
            _width = width;
            _attempts = attempts;
            _planeCodes = CreateVp8lPlaneCodes(width);
        }

        // The chains are only needed while searching; distance codes need only the plane table.
        internal void ReleaseChains() {
            _heads = Array.Empty<int>();
            _previous = Array.Empty<int>();
        }

        internal int DistanceCode(int distance) =>
            distance < _planeCodes.Length && _planeCodes[distance] != 0 ? _planeCodes[distance] : distance + 120;

        internal void Insert(int position) {
            if (position + 1 >= _pixels.Length) return;
            int hash = Hash(position);
            _previous[position] = _heads[hash];
            _heads[hash] = position + 1;
        }

        internal int Find(int position, out int distance) {
            uint[] pixels = _pixels;
            distance = 0;
            int maximum = Math.Min(Vp8lMaximumCopyLength, pixels.Length - position);
            if (maximum < 2) return 0;
            int best = 0;
            // Repeating the previous pixel or row codes in a handful of bits.
            if (position >= 1) Try(position, 1, maximum, ref best, ref distance);
            if (_width > 1 && position >= _width) Try(position, _width, maximum, ref best, ref distance);
            if (best == maximum) return best;
            int candidate = _heads[Hash(position)] - 1;
            int limit = Math.Max(0, position - Vp8lMaximumCopyDistance);
            for (int attempt = 0; attempt < _attempts && candidate >= limit; attempt++) {
                // A candidate must extend the best match to be worth comparing in full.
                if (pixels[candidate + best] == pixels[position + best] && pixels[candidate] == pixels[position]) {
                    int length = MatchLength(candidate, position, maximum);
                    int offset = position - candidate;
                    if (length > best || length == best && length > 0 && DistanceCode(offset) < DistanceCode(distance)) {
                        best = length;
                        distance = offset;
                        if (best == maximum) break;
                    }
                }
                candidate = _previous[candidate] - 1;
            }
            return best;
        }

        private void Try(int position, int offset, int maximum, ref int best, ref int distance) {
            int length = MatchLength(position - offset, position, maximum);
            if (length > best) {
                best = length;
                distance = offset;
            }
        }

        private int MatchLength(int candidate, int position, int maximum) {
            uint[] pixels = _pixels;
            int length = 0;
            while (length < maximum && pixels[candidate + length] == pixels[position + length]) length++;
            return length;
        }

        private int Hash(int position) {
            uint hash = unchecked(_pixels[position] * 0x1E35A7BDU ^ _pixels[position + 1] * 0x9E3779B9U);
            return (int)(unchecked(hash * 0x85EBCA6BU) >> _hashShift);
        }
    }

    // About one chain head per pixel, up to 2^18.
    private static int GetVp8lMatchHashBits(int pixels) {
        int bits = 8;
        while (bits < Vp8lMaximumMatchHashBits && 1 << bits < pixels) bits++;
        return bits;
    }

    // The 120 short codes name nearby (dx, dy) offsets; each distance keeps its smallest code.
    private static int[] CreateVp8lPlaneCodes(int width) {
        var codes = new int[checked(width * 8 + 16)];
        for (int code = 120; code >= 1; code--) {
            int distance = MapVp8lDistance(code, width);
            if (distance > 0 && distance < codes.Length) codes[distance] = code;
        }
        return codes;
    }

    /// <summary>Greedy parse: the longest match wherever it is at least three pixels, or two for the cheap distances.</summary>
    private static Vp8lTokens CreateVp8lTokens(uint[] pixels, int width, int attempts, CancellationToken cancellationToken) {
        var matcher = new Vp8lMatcher(pixels, width, attempts);
        var tokens = new Vp8lTokens(Math.Min(pixels.Length, 1 << 16));
        int position = 0, checkpoint = 0;
        while (position < pixels.Length) {
            // Copies step over many pixels at once, so the check runs whenever a boundary is passed.
            if (position >= checkpoint) {
                cancellationToken.ThrowIfCancellationRequested();
                checkpoint = position + 1024;
            }
            int length = matcher.Find(position, out int distance);
            if (length >= 3 || length == 2 && (distance == 1 || distance == width)) {
                tokens.AddCopy(length, matcher.DistanceCode(distance));
                for (int index = 0; index < length; index++) matcher.Insert(position + index);
                position += length;
            } else {
                tokens.AddLiteral(pixels[position]);
                matcher.Insert(position);
                position++;
            }
        }
        return tokens;
    }

    private static void GetVp8lPrefix(
        int value,
        int prefixCount,
        out int prefix,
        out int extraBits,
        out int extraValue) {
        if (value <= 4) {
            prefix = value - 1;
            extraBits = 0;
            extraValue = 0;
            return;
        }
        // Prefix p >= 4 covers ((2 + (p & 1)) << e) + 1 onwards with e = (p - 2) >> 1 extra bits.
        int highest = 31 - CountLeadingZeros((uint)(value - 1));
        int second = ((value - 1) >> (highest - 1)) & 1;
        prefix = 2 * highest + second;
        if (prefix >= prefixCount) throw new ArgumentOutOfRangeException(nameof(value));
        extraBits = highest - 1;
        extraValue = (value - 1) & ((1 << extraBits) - 1);
    }

    private static int CountLeadingZeros(uint value) {
#if NET8_0_OR_GREATER
        return System.Numerics.BitOperations.LeadingZeroCount(value);
#else
        if (value == 0) return 32;
        int count = 0;
        while ((value & 0x80000000U) == 0) {
            value <<= 1;
            count++;
        }
        return count;
#endif
    }
}
