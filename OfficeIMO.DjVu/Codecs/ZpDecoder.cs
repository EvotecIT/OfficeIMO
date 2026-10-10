namespace OfficeIMO.DjVu;

// DjVu v3, appendix 3. This follows the arithmetic interval definition rather than
// a reference library's buffered/fast-path decoder. Contexts belong to the calling codec.
internal sealed class ZpDecoder {
    private readonly byte[] _source;
    private readonly int _end;
    private readonly CancellationToken _cancellation;
    private int _position;
    private int _inputByte;
    private int _inputBits;
    private int _interval;
    private int _code;
    private int _decisions;

    internal ZpDecoder(byte[] source, int offset, int length, CancellationToken cancellation) {
        if (offset < 0 || length < 0 || offset > source.Length - length) throw new ArgumentOutOfRangeException(nameof(offset));
        _source = source;
        _position = offset;
        _end = offset + length;
        _cancellation = cancellation;
        _code = ReadCodeByte() << 8 | ReadCodeByte();
    }

    internal int Bit(ref byte context) {
        Tick();
        int state = context;
        int split = _interval + ZpTables.Probability[state];
        // A non-renormalizing MPS advances the interval without adapting its context.
        // The complete interval rule is needed for compatibility with existing streams.
        if (split <= Math.Min(_code, 0x7fff)) {
            _interval = split;
            return state & 1;
        }
        split = Math.Min(split, 0x6000 + ((split + _interval) >> 2));
        int bit = state & 1;
        if (_code >= split) {
            if (_interval >= ZpTables.Threshold[state]) context = ZpTables.MoreProbable[state];
            _interval = split;
        } else {
            bit ^= 1;
            int adjustment = 0x10000 - split;
            _interval += adjustment;
            _code += adjustment;
            context = ZpTables.LessProbable[state];
        }
        Normalize();
        return bit;
    }

    internal int RawBit() {
        Tick();
        return Unconditioned(0x8000 + (_interval >> 1));
    }

    internal int WaveletBit() {
        Tick();
        return Unconditioned(0x8000 + ((_interval * 3) >> 3));
    }

    private int Unconditioned(int split) {
        int bit = 0;
        if (_code >= split) {
            _interval = split;
        } else {
            bit = 1;
            int adjustment = 0x10000 - split;
            _interval += adjustment;
            _code += adjustment;
        }
        Normalize();
        return bit;
    }

    internal int Raw(int bits) {
        int value = 0;
        for (int i = 0; i < bits; i++) value = value << 1 | RawBit();
        return value;
    }

    private void Normalize() {
        while (_interval >= 0x8000) {
            _interval = (_interval << 1) & 0xffff;
            _code = ((_code << 1) | ReadCodeBit()) & 0xffff;
        }
    }

    private int ReadCodeByte() => _position < _end ? _source[_position++] : 255;

    private int ReadCodeBit() {
        if (_inputBits == 0) {
            _inputByte = ReadCodeByte();
            _inputBits = 8;
        }
        return (_inputByte >> --_inputBits) & 1;
    }

    private void Tick() {
        if ((_decisions++ & 4095) == 0) _cancellation.ThrowIfCancellationRequested();
    }
}
