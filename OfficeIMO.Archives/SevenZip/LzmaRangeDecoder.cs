namespace OfficeIMO.Archives.SevenZip;

/// <summary>Binary arithmetic decoder for one bounded LZMA range-coded segment.</summary>
internal sealed class LzmaRangeDecoder {
    private const int ProbabilityScale = 2048;
    private const uint MinimumRange = 1u << 24;
    private readonly Stream _source;
    private long _remaining;
    private uint _interval = uint.MaxValue;
    private uint _value;

    internal LzmaRangeDecoder(Stream source, long size) {
        _source = source;
        _remaining = size;
        if (ReadOctet() != 0) throw new InvalidDataException("The LZMA range stream has an invalid prefix.");
        for (int i = 0; i < 4; i++) _value = (_value << 8) | ReadOctet();
        if (_value == _interval) throw new InvalidDataException("The LZMA range stream has an invalid initial value.");
    }

    internal bool Finished => _remaining == 0 && _value == 0;

    /// <summary>True when the coder is flushed, as it is after an end marker.</summary>
    internal bool CodeIsZero => _value == 0;

    internal int Bit(ushort[] model, int index) {
        int weight = model[index];
        uint split = (_interval / ProbabilityScale) * (uint)weight;
        if (_value < split) {
            model[index] = (ushort)(weight + (ProbabilityScale - weight) / 32);
            _interval = split;
            Refill();
            return 0;
        }
        model[index] = (ushort)(weight - weight / 32);
        _value -= split;
        _interval -= split;
        Refill();
        return 1;
    }

    internal uint Bypass(int count) {
        uint value = 0;
        for (int i = 0; i < count; i++) {
            _interval /= 2;
            value <<= 1;
            if (_value >= _interval) {
                value |= 1;
                _value -= _interval;
            }
            Refill();
        }
        return value;
    }

    internal int Tree(ushort[] model, int start, int bits) {
        int node = 1;
        for (int i = 0; i < bits; i++) node = 2 * node + Bit(model, start + node);
        return node - (1 << bits);
    }

    internal uint ReversedTree(ushort[] model, int start, int bits) {
        int node = 1;
        uint result = 0;
        for (int i = 0; i < bits; i++) {
            int bit = Bit(model, start + node);
            result |= (uint)bit << i;
            node = 2 * node + bit;
        }
        return result;
    }

    private void Refill() {
        if (_interval >= MinimumRange) return;
        _interval <<= 8;
        _value = (_value << 8) | ReadOctet();
    }

    private byte ReadOctet() {
        if (_remaining <= 0) throw new InvalidDataException("The LZMA range stream ended inside a symbol.");
        int value = _source.ReadByte();
        if (value < 0) throw new InvalidDataException("The LZMA range stream is truncated.");
        _remaining--;
        return (byte)value;
    }
}
