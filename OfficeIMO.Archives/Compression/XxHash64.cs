namespace OfficeIMO.Archives.Compression;

/// <summary>Incremental XXH64 as defined by the xxHash specification.</summary>
internal sealed class XxHash64 {
    private const ulong Prime1 = 11400714785074694791UL;
    private const ulong Prime2 = 14029467366897019727UL;
    private const ulong Prime3 = 1609587929392839161UL;
    private const ulong Prime4 = 9650029242287828579UL;
    private const ulong Prime5 = 2870177450012600261UL;

    private readonly ulong _seed;
    private readonly byte[] _pending = new byte[32];
    private ulong _v1, _v2, _v3, _v4;
    private int _pendingLength;
    private ulong _length;

    internal XxHash64(ulong seed = 0) {
        _seed = seed;
        Reset();
    }

    internal void Reset() {
        _v1 = unchecked(_seed + Prime1 + Prime2);
        _v2 = unchecked(_seed + Prime2);
        _v3 = _seed;
        _v4 = unchecked(_seed - Prime1);
        _pendingLength = 0;
        _length = 0;
    }

    internal static ulong Compute(byte[] data, int offset, int count, ulong seed = 0) {
        var hash = new XxHash64(seed);
        hash.Append(data, offset, count);
        return hash.Complete();
    }

    internal void Append(byte[] data, int offset, int count) {
        _length += (ulong)count;
        if (_pendingLength > 0) {
            int take = Math.Min(32 - _pendingLength, count);
            Buffer.BlockCopy(data, offset, _pending, _pendingLength, take);
            _pendingLength += take;
            offset += take;
            count -= take;
            if (_pendingLength < 32) return;
            Stripe(_pending, 0);
            _pendingLength = 0;
        }
        while (count >= 32) {
            Stripe(data, offset);
            offset += 32;
            count -= 32;
        }
        if (count > 0) {
            Buffer.BlockCopy(data, offset, _pending, 0, count);
            _pendingLength = count;
        }
    }

    internal ulong Complete() {
        ulong hash;
        unchecked {
            if (_length >= 32) {
                hash = RotateLeft(_v1, 1) + RotateLeft(_v2, 7) + RotateLeft(_v3, 12) + RotateLeft(_v4, 18);
                hash = Merge(hash, _v1);
                hash = Merge(hash, _v2);
                hash = Merge(hash, _v3);
                hash = Merge(hash, _v4);
            } else {
                hash = _seed + Prime5;
            }
            hash += _length;
            int index = 0;
            for (; index + 8 <= _pendingLength; index += 8) {
                hash ^= Round(0, Read64(_pending, index));
                hash = RotateLeft(hash, 27) * Prime1 + Prime4;
            }
            if (index + 4 <= _pendingLength) {
                hash ^= Read32(_pending, index) * Prime1;
                hash = RotateLeft(hash, 23) * Prime2 + Prime3;
                index += 4;
            }
            for (; index < _pendingLength; index++) {
                hash ^= _pending[index] * Prime5;
                hash = RotateLeft(hash, 11) * Prime1;
            }
            hash ^= hash >> 33;
            hash *= Prime2;
            hash ^= hash >> 29;
            hash *= Prime3;
            hash ^= hash >> 32;
        }
        return hash;
    }

    private void Stripe(byte[] data, int offset) {
        _v1 = Round(_v1, Read64(data, offset));
        _v2 = Round(_v2, Read64(data, offset + 8));
        _v3 = Round(_v3, Read64(data, offset + 16));
        _v4 = Round(_v4, Read64(data, offset + 24));
    }

    private static ulong Round(ulong accumulator, ulong lane) =>
        unchecked(RotateLeft(accumulator + lane * Prime2, 31) * Prime1);

    private static ulong Merge(ulong hash, ulong accumulator) =>
        unchecked((hash ^ Round(0, accumulator)) * Prime1 + Prime4);

    private static ulong RotateLeft(ulong value, int count) => (value << count) | (value >> (64 - count));

    private static ulong Read64(byte[] data, int offset) =>
        Read32(data, offset) | Read32(data, offset + 4) << 32;

    private static ulong Read32(byte[] data, int offset) =>
        (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);
}
