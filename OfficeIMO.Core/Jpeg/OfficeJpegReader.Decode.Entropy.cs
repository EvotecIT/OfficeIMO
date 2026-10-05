using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private struct HuffmanTable {
        public int[] Left;
        public int[] Right;
        public int[] Symbols;
        public short[]? Fast;
        public bool IsValid;

        public static HuffmanTable Build(OfficeByteView counts, byte[] values) {
            var left = new int[512];
            var right = new int[512];
            var symbols = new int[512];
            var fast = new short[1 << HuffmanFastBits];
            for (var i = 0; i < left.Length; i++) left[i] = -1;
            for (var i = 0; i < right.Length; i++) right[i] = -1;
            for (var i = 0; i < symbols.Length; i++) symbols[i] = -1;
            for (var i = 0; i < fast.Length; i++) fast[i] = -1;

            var next = 1;
            var code = 0;
            var k = 0;
            for (var i = 1; i <= 16; i++) {
                var count = counts[i - 1];
                for (var j = 0; j < count; j++) {
                    var symbol = values[k++];
                    var node = 0;
                    for (var bit = i - 1; bit >= 0; bit--) {
                        var b = (code >> bit) & 1;
                        if (b == 0) {
                            if (left[node] < 0) {
                                if (next >= left.Length) throw new FormatException("Invalid JPEG Huffman tree.");
                                left[node] = next++;
                            }
                            node = left[node];
                        } else {
                            if (right[node] < 0) {
                                if (next >= right.Length) throw new FormatException("Invalid JPEG Huffman tree.");
                                right[node] = next++;
                            }
                            node = right[node];
                        }
                    }
                    symbols[node] = symbol;
                    if (i <= HuffmanFastBits) {
                        var fill = 1 << (HuffmanFastBits - i);
                        var start = code << (HuffmanFastBits - i);
                        var entry = (short)((i << 8) | symbol);
                        for (var f = 0; f < fill; f++) {
                            fast[start + f] = entry;
                        }
                    }
                    code++;
                }
                code <<= 1;
            }

            return new HuffmanTable {
                Left = left,
                Right = right,
                Symbols = symbols,
                Fast = fast,
                IsValid = true
            };
        }
    }

    private ref struct JpegBitReader {
        private readonly OfficeByteView _data;
        private readonly bool _allowTruncated;
        private readonly CancellationToken _cancellationToken;
        private int _pos;
        private int _bitBuffer;
        private int _bitCount;

        public bool RestartMarkerSeen;

        public JpegBitReader(
            OfficeByteView data,
            bool allowTruncated,
            CancellationToken cancellationToken) {
            _data = data;
            _allowTruncated = allowTruncated;
            _cancellationToken = cancellationToken;
            _pos = 0;
            _bitBuffer = 0;
            _bitCount = 0;
            RestartMarkerSeen = false;
        }

        public bool AllowTruncated => _allowTruncated;

        public bool TryPeekBits(int count, out int value) {
            int originalPosition = _pos;
            int originalBitBuffer = _bitBuffer;
            int originalBitCount = _bitCount;
            bool originalRestartMarkerSeen = RestartMarkerSeen;
            while (_bitCount < count) {
                if (!TryReadByte(out int next)) {
                    RestorePeekState(originalPosition, originalBitBuffer, originalBitCount, originalRestartMarkerSeen);
                    value = 0;
                    return false;
                }
                if (RestartMarkerSeen != originalRestartMarkerSeen) {
                    RestorePeekState(originalPosition, originalBitBuffer, originalBitCount, originalRestartMarkerSeen);
                    value = 0;
                    return false;
                }
                _bitBuffer = (_bitBuffer << 8) | next;
                _bitCount += 8;
            }
            value = (_bitBuffer >> (_bitCount - count)) & ((1 << count) - 1);
            return true;
        }

        private void RestorePeekState(int position, int bitBuffer, int bitCount, bool restartMarkerSeen) {
            _pos = position;
            _bitBuffer = bitBuffer;
            _bitCount = bitCount;
            RestartMarkerSeen = restartMarkerSeen;
        }

        public void SkipBits(int count) {
            if (count == 0) return;
            _bitCount -= count;
            if (_bitCount <= 0) {
                _bitCount = 0;
                _bitBuffer = 0;
            } else {
                _bitBuffer &= (1 << _bitCount) - 1;
            }
        }

        public int ReadBit() {
            EnsureBits(1);
            var bit = (_bitBuffer >> (_bitCount - 1)) & 1;
            _bitCount--;
            if (_bitCount == 0) {
                _bitBuffer = 0;
            } else {
                _bitBuffer &= (1 << _bitCount) - 1;
            }
            return bit;
        }

        public int ReadBits(int count) {
            if (count == 0) return 0;
            EnsureBits(count);
            var value = (_bitBuffer >> (_bitCount - count)) & ((1 << count) - 1);
            _bitCount -= count;
            if (_bitCount == 0) {
                _bitBuffer = 0;
            } else {
                _bitBuffer &= (1 << _bitCount) - 1;
            }
            return value;
        }


        public void ExpectRestartMarker(int expectedMarker = -1) {
            _bitBuffer = 0;
            _bitCount = 0;
            while (_pos < _data.Length) {
                var b = _data[_pos++];
                if (b != 0xFF) {
                    if (expectedMarker >= 0) throw new FormatException("Unexpected JPEG restart data.");
                    continue;
                }
                SkipFillBytes(_data, ref _pos, _cancellationToken);
                if (_pos >= _data.Length) throw new FormatException("Unexpected JPEG end.");
                var marker = _data[_pos++];
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

        private void EnsureBits(int count) {
            while (_bitCount < count) {
                var b = ReadByte();
                _bitBuffer = (_bitBuffer << 8) | b;
                _bitCount += 8;
            }
        }

        private int ReadByte() {
            if (TryReadByte(out int value)) return value;
            throw new FormatException("Unexpected JPEG end.");
        }

        private bool TryReadByte(out int value) {
            while (_pos < _data.Length) {
                var b = _data[_pos++];
                if (b != 0xFF) {
                    value = b;
                    return true;
                }
                SkipFillBytes(_data, ref _pos, _cancellationToken);
                if (_pos >= _data.Length) {
                    if (_allowTruncated) {
                        value = 0;
                        return true;
                    }
                    value = 0;
                    return false;
                }
                var marker = _data[_pos++];
                if (marker == 0x00) {
                    value = 0xFF;
                    return true;
                }
                if (marker >= 0xD0 && marker <= 0xD7) {
                    RestartMarkerSeen = true;
                    continue;
                }
                throw new FormatException("Unexpected JPEG marker in scan.");
            }
            if (_allowTruncated) {
                value = 0;
                return true;
            }
            value = 0;
            return false;
        }
    }

}
