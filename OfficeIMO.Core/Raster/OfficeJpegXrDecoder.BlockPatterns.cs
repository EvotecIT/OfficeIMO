using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private sealed class BlockPatterns {
        private readonly int _components, _color;
        private readonly BinaryVlc _groups = new(), _blocks = new();
        private readonly int[] _state = new int[2], _ones = { -4, -4 }, _zeroes = { 4, 4 };
        private readonly int[] _difference;
        internal BlockPatterns(int components, int color = 0) { _components = components; _color = color; _difference = new int[components]; }

        // T.832 8.7.17.2/8.7.17.5 and 8.10, for gray, YUV and YUVK planes.
        internal void Read(Bits bits, int[] patterns, int start, int left, int top, bool leftEdge, bool topEdge) {
            Array.Clear(_difference, 0, _difference.Length);
            bool independent = _color == 0 || _color == 4 || _color == 6;
            for (int channel = 0; channel < (independent ? _components : 1); channel++) {
                int groupMask = Refine(bits, _groups.Read(bits, PatternCountCodes, PatternCountDeltas));
                for (int group = 0; group < 4; group++) if ((groupMask & (1 << group)) != 0) {
                    int value = independent ? _blocks.Read(bits, PatternCountCodes, PatternCountDeltas) + 1
                        : _blocks.Read(bits, ChromaPatternCountCodes, ChromaPatternCountDeltas) + 1;
                    int chroma = 0;
                    if (value >= 6) {
                        chroma = (TernaryCodes.Read(bits) + 1) << 4;
                        if (value >= 9) value += TernaryCodes.Read(bits);
                        value -= 6;
                    }
                    if (value > 5) throw new FormatException("JPEG-XR block pattern symbol is invalid.");
                    int index = PatternOffsets[value] + (int)bits.Read(PatternWidths[value]);
                    int pattern = PatternValues[index] | chroma;
                    _difference[channel] |= (pattern & 15) << (group * 4);
                    for (int c = 1; !independent && c < _components; c++) if ((pattern & (1 << (c + 3))) != 0) {
                        if (_color == 1) _difference[c] |= 1 << group;
                        else if (_color == 2) {
                            int symbol = TernaryCodes.Read(bits);
                            int pair = symbol == 0 ? 1 : symbol == 1 ? 4 : 5;
                            int shift = group < 2 ? group : group + 2;
                            _difference[c] |= pair << shift;
                        } else _difference[c] |= Refine(bits, ChromaBlockCodes.Read(bits) + 1) << (group * 4);
                    }
                }
            }
            for (int c = 0; c < _components; c++) {
                int model = c == 0 ? 0 : 1, pattern = _difference[c];
                bool reduced = c != 0 && (_color == 1 || _color == 2);
                if (reduced) {
                    if (_state[model] == 0) {
                        pattern ^= leftEdge ? topEdge ? 1 : (patterns[top + c] >> (_color == 1 ? 2 : 6)) & 1
                            : (patterns[left + c] >> 1) & 1;
                        pattern ^= (pattern & 1) << 1;
                        pattern ^= (pattern & 3) << 2;
                        if (_color == 2) { pattern ^= (pattern & 12) << 2; pattern ^= (pattern & 48) << 2; }
                    } else if (_state[model] == 2) pattern ^= _color == 1 ? 15 : 255;
                } else {
                    if (_state[model] == 0) {
                        pattern ^= leftEdge ? topEdge ? 1 : (patterns[top + c] >> 10) & 1 : (patterns[left + c] >> 5) & 1;
                        pattern ^= 2 & (pattern << 1);
                        pattern ^= 0x10 & (pattern << 3);
                        pattern ^= 0x20 & (pattern << 1);
                        pattern ^= (pattern & 0x33) << 2;
                        pattern ^= (pattern & 0xCC) << 6;
                        pattern ^= (pattern & 0x3300) << 2;
                    } else if (_state[model] == 2) pattern ^= 0xFFFF;
                }
                patterns[start + c] = pattern;
                int ones = 0;
                for (int value = pattern; value != 0; value &= value - 1) ones++;
                if (reduced) ones *= _color == 1 ? 4 : 2;
                _ones[model] = Math.Max(-16, Math.Min(15, _ones[model] + ones - 3));
                _zeroes[model] = Math.Max(-16, Math.Min(15, _zeroes[model] + 13 - ones));
                _state[model] = _ones[model] < 0 ? _ones[model] < _zeroes[model] ? 1 : 2 : _zeroes[model] < 0 ? 2 : 0;
            }
        }

        internal void Adapt() { _groups.Adapt(); _blocks.Adapt(); }
        private static int Refine(Bits bits, int count) => count == 2 ? DoublePatterns[DoublePatternCodes.Read(bits)]
            : count == 1 ? 1 << (int)bits.Read(2) : count == 3 ? 15 ^ (1 << (int)bits.Read(2)) : count == 4 ? 15 : 0;
    }
}
