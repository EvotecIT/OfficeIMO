using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // T.81 D.2, with Cx in the high word and incoming bits in the low word.
    // Arithmetic termination omits trailing zero bytes; a marker supplies zeros
    // until the current MCU interval is complete, without consuming that marker.
    private sealed class ArithmeticReader {
        private readonly OfficeByteView _data;
        private readonly CancellationToken _token;
        private int _offset, _bits, _restart;
        private uint _interval, _code;

        public ArithmeticReader(OfficeByteView data, CancellationToken token) {
            _data = data; _token = token; Initialize();
        }

        private void Initialize() {
            _interval = 0x10000;
            _code = ((uint)ReadByte() << 24) | ((uint)ReadByte() << 16);
            _bits = 0;
        }

        private int ReadByte() {
            if ((_offset & 4095) == 0) _token.ThrowIfCancellationRequested();
            if (_offset == _data.Length) return 0;
            int value = _data[_offset];
            if (value != 255) { _offset++; return value; }
            if (_offset + 1 < _data.Length && _data[_offset + 1] == 0) { _offset += 2; return 255; }
            return 0;
        }

        public int Decode(ref byte context) {
            int state = context & 127, mps = context >> 7;
            uint estimate = ArithmeticEstimates[state], qe = estimate >> 16;
            _interval -= qe;
            bool upper = (_code >> 16) >= _interval;
            if (!upper && _interval >= 0x8000) return mps;
            bool lps = upper ? _interval >= qe : _interval < qe;
            if (upper) { _code -= _interval << 16; _interval = qe; }
            int decision = lps ? 1 - mps : mps;
            context = lps ? (byte)((estimate >> 8) ^ (mps << 7))
                : (byte)((estimate & 255) | ((uint)mps << 7));
            do {
                if (_bits == 0) { _code |= (uint)ReadByte() << 8; _bits = 8; }
                _code <<= 1; _interval <<= 1; _bits--;
            } while (_interval < 0x8000);
            return decision;
        }

        public int DecodeFixed() { byte context = 0; return Decode(ref context); }

        public void Finish() {
            while (_offset < _data.Length) {
                _token.ThrowIfCancellationRequested();
                if (_data[_offset++] != 255) continue;
                if (_offset == _data.Length || _data[_offset++] != 0)
                    throw new FormatException("Unexpected arithmetic JPEG marker after the final MCU.");
            }
        }

        public void Restart() {
            // The final interval may leave encoded bytes unread. Consume only to
            // its explicit restart marker, respecting FF00 and marker fill bytes.
            while (_offset < _data.Length) {
                _token.ThrowIfCancellationRequested();
                if (_data[_offset++] != 255) continue;
                while (_offset < _data.Length && _data[_offset] == 255) {
                    if ((_offset & 4095) == 0) _token.ThrowIfCancellationRequested();
                    _offset++;
                }
                if (_offset == _data.Length) break;
                int marker = _data[_offset++];
                if (marker == 0) continue;
                if (marker != 0xD0 + _restart) throw new FormatException("Unexpected arithmetic JPEG restart marker.");
                _restart = (_restart + 1) & 7; Initialize(); return;
            }
            throw new FormatException("Missing arithmetic JPEG restart marker.");
        }
    }
}
