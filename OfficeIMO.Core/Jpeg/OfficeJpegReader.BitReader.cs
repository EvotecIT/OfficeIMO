#if NET8_0_OR_GREATER
using System.Buffers.Binary;
#endif
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // Entropy-coded segment reader. Bytes are prefetched into a 64-bit buffer, but the
    // prefetch never crosses a marker or the end of the data: a restart marker is
    // consumed (and RestartMarkerSeen raised) only when a read needs bits beyond it,
    // and a missing byte is reported only when a read needs it. Readers created
    // without prefetch load bytes exactly as they are needed, which keeps the position
    // seen by ExpectRestartMarker identical to a byte-at-a-time reader.
    private ref struct JpegBitReader {
        private const int PrefetchBits = 57;
        private readonly byte[] _bytes;
        private readonly int _end;
        private readonly bool _allowTruncated;
        private readonly bool _prefetch;
        private readonly CancellationToken _cancellationToken;
        private int _position;
        private ulong _buffer;
        private int _bitCount;

        // Hot loops copy the bit buffer into locals and hand it back before any other read.
        public readonly ulong BufferedBits => _buffer;
        public int BufferedBitCount {
            readonly get => _bitCount;
            set => _bitCount = value;
        }

        public bool RestartMarkerSeen;

        public JpegBitReader(
            OfficeByteView data,
            bool allowTruncated,
            CancellationToken cancellationToken,
            bool prefetch = true) {
            _bytes = data.UnderlyingArray;
            _position = data.Offset;
            _end = data.Offset + data.Length;
            _allowTruncated = allowTruncated;
            _prefetch = prefetch;
            _cancellationToken = cancellationToken;
            _buffer = 0;
            _bitCount = 0;
            RestartMarkerSeen = false;
        }

        public bool AllowTruncated => _allowTruncated;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool TryPeekBits(int count, out int value) {
            if (_bitCount < count && !TryFillForPeek(count)) {
                value = 0;
                return false;
            }
            value = (int)(_buffer >> (_bitCount - count)) & ((1 << count) - 1);
            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void SkipBits(int count) {
            _bitCount -= count;
            if (_bitCount < 0) _bitCount = 0;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public int ReadBit() {
            if (_bitCount < 1) EnsureBits(1);
            _bitCount--;
            return (int)(_buffer >> _bitCount) & 1;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public int ReadBits(int count) {
            if (count == 0) return 0;
            if (_bitCount < count) EnsureBits(count);
            _bitCount -= count;
            return (int)(_buffer >> _bitCount) & ((1 << count) - 1);
        }

        public void ExpectRestartMarker(int expectedMarker = -1) {
            _buffer = 0;
            _bitCount = 0;
            while (_position < _end) {
                var b = _bytes[_position++];
                if (b != 0xFF) {
                    if (expectedMarker >= 0) throw new FormatException("Unexpected JPEG restart data.");
                    continue;
                }
                SkipFill(ref _position);
                if (_position >= _end) throw new FormatException("Unexpected JPEG end.");
                var marker = _bytes[_position++];
                if (marker >= 0xD0 && marker <= 0xD7) {
                    if (expectedMarker >= 0 && marker != expectedMarker) throw new FormatException("Unexpected JPEG restart sequence.");
                    RestartMarkerSeen = false;
                    return;
                }
                if (marker == 0x00 && expectedMarker < 0) continue;
                throw new FormatException("Unexpected JPEG marker in scan.");
            }
            throw new FormatException("Missing JPEG restart marker.");
        }

        private bool TryFillForPeek(int count) {
            if (_prefetch) {
                Fill(PrefetchBits);
                return _bitCount >= count;
            }
            int position = _position;
            ulong buffer = _buffer;
            int bitCount = _bitCount;
            Fill(count);
            if (_bitCount >= count) return true;
            _position = position;
            _buffer = buffer;
            _bitCount = bitCount;
            return false;
        }

        public void EnsureBits(int count) {
            Fill(_prefetch ? PrefetchBits : count);
            while (_bitCount < count) {
                // Fill stops only at the end of the data or at a marker.
                if (_position >= _end) throw new FormatException("Unexpected JPEG end.");
                int next = _position + 1;
                SkipFill(ref next);
                var marker = _bytes[next];
                if (marker < 0xD0 || marker > 0xD7) throw new FormatException("Unexpected JPEG marker in scan.");
                RestartMarkerSeen = true;
                _position = next + 1;
                Fill(_prefetch ? PrefetchBits : count);
            }
        }

        // Loads whole bytes until the buffer holds at least target bits, undoing byte
        // stuffing. Stops before a marker, and at the end of non-truncated data.
        private void Fill(int target) {
            // Whole-word refill when the next eight bytes hold no 0xFF (no stuffing or marker).
            if (_bitCount < target && _end - _position >= 8) {
#if NET8_0_OR_GREATER
                ulong word = BinaryPrimitives.ReadUInt64BigEndian(new ReadOnlySpan<byte>(_bytes, _position, 8));
#else
                ulong word = 0;
                for (int i = 0; i < 8; i++) word = (word << 8) | _bytes[_position + i];
#endif
                ulong inverted = ~word;
                if (((inverted - 0x0101010101010101UL) & ~inverted & 0x8080808080808080UL) == 0) {
                    int bytes = (target - _bitCount + 7) >> 3;
                    if (bytes >= 8) {
                        _buffer = word;
                        bytes = 8;
                    } else {
                        _buffer = (_buffer << (bytes * 8)) | (word >> (64 - bytes * 8));
                    }
                    _bitCount += bytes * 8;
                    _position += bytes;
                    return;
                }
            }
            while (_bitCount < target) {
                if (_position >= _end) {
                    if (!_allowTruncated) return;
                    _buffer <<= 8;
                    _bitCount += 8;
                    continue;
                }
                byte value = _bytes[_position];
                if (value != 0xFF) {
                    _buffer = (_buffer << 8) | value;
                    _bitCount += 8;
                    _position++;
                    continue;
                }
                int next = _position + 1;
                SkipFill(ref next);
                if (next >= _end) {
                    // Trailing fill bytes end the data.
                    _position = _end;
                    continue;
                }
                if (_bytes[next] != 0x00) return;
                _buffer = (_buffer << 8) | 0xFF;
                _bitCount += 8;
                _position = next + 1;
            }
        }

        private void SkipFill(ref int offset) {
            int fillBytes = 0;
            while (offset < _end && _bytes[offset] == 0xFF) {
                if ((fillBytes++ & 0xFFF) == 0) _cancellationToken.ThrowIfCancellationRequested();
                offset++;
            }
        }
    }
}
