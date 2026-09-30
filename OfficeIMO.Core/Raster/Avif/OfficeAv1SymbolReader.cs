using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>AV1 Q15 arithmetic symbol decoder bounded to one tile and a caller-supplied symbol budget.</summary>
/// <remarks>Implements AV1 sections 8.2.2–8.2.6. Syntax/context selection and pixel reconstruction are separate owners.</remarks>
internal sealed class OfficeAv1SymbolReader {
    private readonly byte[] _bytes;
    private readonly int _offset;
    private readonly int _lengthBits;
    private readonly bool _disableCdfUpdate;
    private readonly CancellationToken _cancellation;
    private long _symbolsRemaining;
    private int _position;
    private int _range = 32768;
    private int _value;
    private int _remainingBits;
    private bool _finished;
    private bool _failed;

    internal OfficeAv1SymbolReader(byte[] bytes, int offset, int length, long maximumSymbols,
        bool disableCdfUpdate, CancellationToken cancellation = default) {
        cancellation.ThrowIfCancellationRequested();
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (offset < 0 || length < 1 || offset > bytes.Length - length || length > 128 * 1024 * 1024)
            throw new FormatException("Invalid AV1 entropy payload bounds.");
        if (maximumSymbols < 1) throw new ArgumentOutOfRangeException(nameof(maximumSymbols));
        _bytes = bytes; _offset = offset; _lengthBits = checked(length * 8);
        _disableCdfUpdate = disableCdfUpdate; _cancellation = cancellation; _symbolsRemaining = maximumSymbols;
        int count = Math.Min(_lengthBits, 15);
        _value = 32767 ^ (ReadBits(count) << (15 - count));
        _remainingBits = _lengthBits - 15;
    }

    /// <summary>Reads and optionally adapts a rising CDF of N probabilities followed by its saturating update count.</summary>
    internal int ReadSymbol(int[] cdf) {
        if (cdf == null) throw new ArgumentNullException(nameof(cdf));
        int count = cdf.Length - 1;
        Require(count >= 2 && count <= 16 && cdf[0] > 0 && cdf[count - 1] == 32768 && cdf[count] >= 0 && cdf[count] <= 32);
        int previous = 0;
        for (int i = 0; i < count; i++) {
            Require(cdf[i] >= previous && cdf[i] <= 32768);
            previous = cdf[i];
        }
        return ReadCore(cdf, count);
    }

    /// <summary>Reads a pseudo-raw bit without retaining the temporary equal-probability CDF.</summary>
    internal bool ReadBool() => ReadCore(null, 2) != 0;

    internal int ReadLiteral(int count) {
        if (count < 0 || count > 31) throw new ArgumentOutOfRangeException(nameof(count));
        int value = 0;
        for (int i = 0; i < count; i++) value = (value << 1) | (ReadBool() ? 1 : 0);
        return value;
    }

    private int ReadCore(int[]? cdf, int count) {
        _cancellation.ThrowIfCancellationRequested();
        Require(!_finished && !_failed && _symbolsRemaining > 0);
        _symbolsRemaining--;
        int upper = _range, lower = _range, symbol = -1;
        do {
            upper = lower;
            symbol++;
            int probability = cdf == null ? (symbol == 0 ? 16384 : 32768) : cdf[symbol];
            lower = ((_range >> 8) * ((32768 - probability) >> 6)) >> 1;
            lower += 4 * (count - symbol - 1);
        } while (_value < lower);
        Require(symbol < count && lower < upper && upper <= _range);
        _range = upper - lower;
        _value -= lower;
        int shift = 0;
        while ((_range << shift) < 32768) shift++;
        _range <<= shift;
        int available = Math.Min(shift, Math.Max(0, _remainingBits));
        int bits = ReadBits(available) << (shift - available);
        _value = bits ^ (((_value + 1) << shift) - 1);
        _remainingBits -= shift;
        // Once more than fourteen synthetic zeros were consumed, no later exit can be conforming.
        Require(_remainingBits >= -14);
        if (cdf != null && !_disableCdfUpdate) {
            int rate = 3 + (cdf[count] > 15 ? 1 : 0) + (cdf[count] > 31 ? 1 : 0) + Math.Min(FloorLog2(count), 2);
            for (int i = 0; i < count - 1; i++) {
                int target = i >= symbol ? 32768 : 0;
                if (target < cdf[i]) cdf[i] -= (cdf[i] - target) >> rate;
                else cdf[i] += (target - cdf[i]) >> rate;
            }
            if (cdf[count] < 32) cdf[count]++;
        }
        return symbol;
    }

    /// <summary>Checks the tile's mandatory trailing one and zero padding after the caller has decoded its complete syntax.</summary>
    internal void Finish() {
        _cancellation.ThrowIfCancellationRequested();
        Require(!_finished && !_failed && _remainingBits >= -14);
        int trailing = _position - Math.Min(15, _remainingBits + 15);
        Require(trailing >= 0 && trailing < _lengthBits && BitAt(trailing) == 1);
        int p = trailing + 1;
        while ((p & 7) != 0 && p < _lengthBits) Require(BitAt(p++) == 0);
        while (p < _lengthBits) {
            if ((p & 8191) == 0) _cancellation.ThrowIfCancellationRequested();
            Require(_bytes[_offset + (p >> 3)] == 0);
            p += 8;
        }
        _finished = true;
    }

    private int ReadBits(int count) {
        Require(count >= 0 && count <= 15 && _position <= _lengthBits - count);
        int value = 0;
        for (int i = 0; i < count; i++) value = (value << 1) | BitAt(_position++);
        return value;
    }

    private int BitAt(int position) => (_bytes[_offset + (position >> 3)] >> (7 - (position & 7))) & 1;
    private static int FloorLog2(int value) { int result = 0; while ((value >>= 1) != 0) result++; return result; }
    private void Require(bool condition) {
        if (!condition) { _failed = true; throw new FormatException("Invalid or exhausted AV1 entropy payload."); }
    }
}
