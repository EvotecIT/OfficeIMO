namespace OfficeIMO.Pdf.Filters;

internal static partial class Jbig2Decoder {
    // MQ arithmetic decoder (T.88 Annex E, the INITDEC, DECODE, RENORMD and BYTEIN procedures).
    private sealed class ArithmeticDecoder {
        private static readonly ushort[] Qe = {
            0x5601, 0x3401, 0x1801, 0x0AC1, 0x0521, 0x0221, 0x5601, 0x5401, 0x4801, 0x3801, 0x3001, 0x2401, 0x1C01, 0x1601,
            0x5601, 0x5401, 0x5101, 0x4801, 0x3801, 0x3401, 0x3001, 0x2801, 0x2401, 0x2201, 0x1C01, 0x1801, 0x1601, 0x1401,
            0x1201, 0x1101, 0x0AC1, 0x09C1, 0x08A1, 0x0521, 0x0441, 0x02A1, 0x0221, 0x0141, 0x0111, 0x0085, 0x0049, 0x0025,
            0x0015, 0x0009, 0x0005, 0x0001, 0x5601
        };
        private static readonly byte[] NextMps = {
            1, 2, 3, 4, 5, 38, 7, 8, 9, 10, 11, 12, 13, 29, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32,
            33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 45, 46
        };
        private static readonly byte[] NextLps = {
            1, 6, 9, 12, 29, 33, 6, 14, 14, 14, 17, 18, 20, 21, 14, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29,
            30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 46
        };
        private static readonly byte[] Switch = {
            1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
            0, 0, 0, 0, 0, 0, 0
        };

        private readonly byte[] _data;
        private readonly int _end;
        private int _position;
        private uint _c;
        private uint _a;
        private int _ct;
        private readonly Action<long> _charge;

        internal ArithmeticDecoder(byte[] data, int start, int end, Action<long> charge) {
            if (start < 0 || start >= end || end > data.Length) throw new InvalidDataException("Truncated JBIG2 arithmetic data.");
            _charge = charge;
            _data = data;
            _end = end;
            _position = start;
            _c = (uint)(ByteAt(_position) ^ 0xFF) << 16;
            ByteIn();
            _c <<= 7;
            _ct -= 7;
            _a = 0x8000;
        }

        private int ByteAt(int position) => position < _end ? _data[position] : 0xFF;

        private void ByteIn() {
            if (ByteAt(_position) == 0xFF) {
                int next = ByteAt(_position + 1);
                if (next > 0x8F) {
                    _ct = 8;
                } else {
                    _position++;
                    _c += 0xFE00u - ((uint)ByteAt(_position) << 9);
                    _ct = 7;
                }
            } else {
                _position++;
                _c += 0xFF00u - ((uint)ByteAt(_position) << 8);
                _ct = 8;
            }
        }

        // Contexts store (index << 1) | mps. DECODE with the inverted code register of T.88 E.3.2: the MPS
        // sub-interval lies below A.
        internal int Decode(byte[] contexts, int context) {
            _charge(1L);
            int state = contexts[context] >> 1;
            int mps = contexts[context] & 1;
            uint qe = Qe[state];
            int decision;
            _a -= qe;
            if ((_c >> 16) < _a) {
                if ((_a & 0x8000) != 0) return mps;
                if (_a < qe) {
                    decision = 1 - mps;
                    if (Switch[state] == 1) mps = 1 - mps;
                    state = NextLps[state];
                } else {
                    decision = mps;
                    state = NextMps[state];
                }
            } else {
                _c -= _a << 16;
                if (_a < qe) {
                    decision = mps;
                    state = NextMps[state];
                } else {
                    decision = 1 - mps;
                    if (Switch[state] == 1) mps = 1 - mps;
                    state = NextLps[state];
                }
                _a = qe;
            }
            do {
                if (_ct == 0) ByteIn();
                _a <<= 1;
                _c <<= 1;
                _ct--;
            } while ((_a & 0x8000) == 0);
            contexts[context] = (byte)((state << 1) | mps);
            return decision;
        }

        // Integer decoding procedure (Annex A.2); null is OOB.
        internal int? DecodeInteger(byte[] contexts) {
            int previous = 1;
            int ReadBits(int count) {
                int value = 0;
                for (int index = 0; index < count; index++) {
                    int bit = Decode(contexts, previous);
                    previous = previous < 256 ? (previous << 1) | bit : (((previous << 1) | bit) & 511) | 256;
                    value = (value << 1) | bit;
                }
                return value;
            }
            int sign = ReadBits(1);
            long magnitude = ReadBits(1) == 0 ? ReadBits(2)
                : ReadBits(1) == 0 ? ReadBits(4) + 4L
                : ReadBits(1) == 0 ? ReadBits(6) + 20L
                : ReadBits(1) == 0 ? ReadBits(8) + 84L
                : ReadBits(1) == 0 ? ReadBits(12) + 340L
                : (uint)ReadBits(32) + 4436L;
            if (sign == 0) return magnitude > int.MaxValue ? throw new InvalidDataException("JBIG2 integer overflow.") : (int)magnitude;
            if (magnitude == 0) return null;
            return magnitude > int.MaxValue ? throw new InvalidDataException("JBIG2 integer overflow.") : -(int)magnitude;
        }

        // IAID decoding procedure (Annex A.3).
        internal int DecodeId(byte[] contexts, int codeLength) {
            int previous = 1;
            for (int index = 0; index < codeLength; index++) previous = (previous << 1) | Decode(contexts, previous);
            return previous - (1 << codeLength);
        }
    }

    private sealed class IntegerContexts {
        private readonly Dictionary<string, byte[]> _contexts = new(StringComparer.Ordinal);
        private readonly Func<int, byte[]> _allocate;
        internal IntegerContexts(Func<int, byte[]> allocate) { _allocate = allocate; }
        internal byte[] this[string name] {
            get {
                if (!_contexts.TryGetValue(name, out byte[]? contexts)) _contexts[name] = contexts = _allocate(512);
                return contexts;
            }
        }
    }

}
