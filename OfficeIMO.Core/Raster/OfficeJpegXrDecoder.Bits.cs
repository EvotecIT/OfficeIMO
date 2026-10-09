using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private sealed class Bits {
        private readonly byte[] _bytes;
        private readonly int _end;
        private readonly CancellationToken _cancellation;
        private int _offset;
        private int _remaining;
        private int _reads;

        internal Bits(byte[] bytes, int offset, int length, CancellationToken cancellation) {
            if (offset < 0 || length < 0 || offset > bytes.Length - length)
                throw new FormatException("JPEG-XR codestream range is invalid.");
            _bytes = bytes; _offset = offset; _end = offset + length; _cancellation = cancellation;
            cancellation.ThrowIfCancellationRequested();
        }

        internal int ByteOffset => _offset;
        internal long BitOffset => (long)_offset * 8 + (_remaining == 0 ? 0 : 8 - _remaining);
        internal uint Read(int count) {
            if (count < 0 || count > 32) throw new ArgumentOutOfRangeException(nameof(count));
            if ((_reads++ & 1023) == 0) _cancellation.ThrowIfCancellationRequested();
            uint value = 0;
            while (count != 0) {
                if (_remaining == 0) {
                    if (_offset == _end) throw new FormatException("JPEG-XR codestream is truncated.");
                    _remaining = 8;
                }
                int take = Math.Min(count, _remaining);
                value = (value << take) | (uint)((_bytes[_offset] >> (_remaining - take)) & ((1 << take) - 1));
                _remaining -= take; count -= take;
                if (_remaining == 0) _offset++;
            }
            return value;
        }
        internal bool Flag() => Read(1) != 0;
        internal void SkipAligned(int count) {
            _cancellation.ThrowIfCancellationRequested();
            if (_remaining != 0 || count < 0 || count > _end - _offset)
                throw new FormatException("JPEG-XR aligned data range is invalid.");
            _offset += count;
        }
        internal void AlignZero() {
            if (_remaining != 0 && Read(_remaining) != 0)
                throw new FormatException("JPEG-XR header alignment is invalid.");
        }
    }
}
