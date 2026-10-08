using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    // Prefix trees keep reads inside the packet, including the final short code.
    private sealed class Codebook {
        private readonly int[,] _edges;
        private readonly int[] _values;

        internal Codebook(params string[] codes) {
            int size = 1;
            foreach (string code in codes) size += code.Length;
            _edges = new int[size, 2]; _values = new int[size];
            for (int i = 0; i < size; i++) _values[i] = -1;
            int next = 1;
            for (int value = 0; value < codes.Length; value++) {
                int node = 0;
                foreach (char bit in codes[value]) {
                    int branch = bit - '0';
                    if (_edges[node, branch] == 0) _edges[node, branch] = next++;
                    node = _edges[node, branch];
                }
                _values[node] = value;
            }
        }

        internal int Read(Bits bits) {
            int node = 0;
            while (_values[node] < 0) {
                node = _edges[node, (int)bits.Read(1)];
                if (node == 0) throw new FormatException("JPEG-XR entropy code is invalid.");
            }
            return _values[node];
        }
    }

    // T.832 Tables 51, 52 and 86.
    private static readonly Codebook DcPresence = new("10", "001", "00001", "0001", "11", "010", "00000", "011");
    private static readonly Codebook[] AbsLevelCodes = {
        new("01", "10", "11", "001", "0001", "00000", "00001"),
        new("1", "01", "001", "0001", "00001", "000000", "000001")
    };
    private static readonly int[] AbsLevelBase = { 2, 3, 4, 6, 10, 14 };
    private static readonly int[] AbsLevelBits = { 0, 0, 1, 2, 2, 2 };

    private sealed class BinaryVlc {
        private int _table, _discriminant;

        internal int Read(Bits bits, Codebook[] codes, int[] deltas) {
            int value = codes[_table].Read(bits); _discriminant += deltas[value]; return value;
        }

        internal int ReadAbsolute(Bits bits) {
            int index = Read(bits, AbsLevelCodes, AbsLevelDeltas);
            if (index < 6) return AbsLevelBase[index] + (int)bits.Read(AbsLevelBits[index]);
            int count = (int)bits.Read(4) + 4;
            if (count == 19) { count += (int)bits.Read(2); if (count == 22) count += (int)bits.Read(3); }
            return checked(2 + (1 << count) + (int)bits.Read(count));
        }

        internal void Adapt() {
            if (_discriminant < -8 && _table != 0) { _table--; _discriminant = 0; }
            else if (_discriminant > 8 && _table != 1) { _table++; _discriminant = 0; }
            else _discriminant = Math.Max(-64, Math.Min(64, _discriminant));
        }
    }

    private sealed class CoefficientModel {
        internal readonly int[] Bits = new int[2];
        private readonly int[] _state = new int[2];
        private readonly int _band, _components, _color;
        private static readonly int[] LumaWeights = { 240, 12, 1 };
        private static readonly int[,] ChromaWeights = { { 0, 240, 120, 80, 60, 48, 40, 34 }, { 0, 12, 6, 4, 3, 2, 2, 2 }, { 0, 16, 8, 5, 4, 3, 3, 2 } };

        internal CoefficientModel(int band, int components, int color = 0) {
            _band = band; _components = components; _color = color;
            Bits[0] = Bits[1] = (2 - band) * 4;
        }

        internal void Update(int lumaCount, int chromaCount) {
            // T.832 Table 112 gives separate normalization weights for reduced chroma planes.
            for (int i = 0; i < (_components == 1 ? 1 : 2); i++) {
                int chromaWeight = _color == 1 ? (_band == 0 ? 120 : _band == 1 ? 37 : 2)
                    : _color == 2 ? (_band == 0 ? 120 : _band == 1 ? 18 : 1)
                    : ChromaWeights[_band, _components - 1];
                int weighted = i == 0 ? lumaCount * LumaWeights[_band] : chromaCount * chromaWeight;
                if (i == 1 && _band == 2 && _color != 1 && _color != 2) weighted >>= 4;
                int delta = (weighted - 70) >> 2;
                if (delta <= -8) {
                    _state[i] += Math.Max(-16, delta + 4);
                    if (_state[i] < -8) {
                        if (Bits[i] == 0) _state[i] = -8;
                        else { _state[i] = 0; Bits[i]--; }
                    }
                } else if (delta >= 8) {
                    _state[i] += Math.Min(15, delta - 4);
                    if (_state[i] > 8) {
                        if (Bits[i] >= 15) { Bits[i] = 15; _state[i] = 8; }
                        else { _state[i] = 0; Bits[i]++; }
                    }
                }
            }
        }
    }

    private sealed class DcContext {
        private readonly BinaryVlc _luma = new(), _chroma = new();
        private readonly CoefficientModel _model;
        private readonly int _components, _color;

        internal DcContext(int components, int color = 0) { _components = components; _color = color; _model = new CoefficientModel(0, components, color); }

        internal void Read(Bits bits, int[] destination, int offset, bool adapt) {
            bool independent = _color == 0 || _color == 4 || _color == 6;
            int presence = independent ? 0 : DcPresence.Read(bits);
            int lumaCount = 0, chromaCount = 0;
            for (int i = 0; i < _components; i++) {
                bool nonzero = independent ? bits.Flag() : (presence & (1 << (_components - i - 1))) != 0;
                if (nonzero) { if (i == 0) lumaCount++; else chromaCount++; }
                BinaryVlc vlc = i == 0 || independent ? _luma : _chroma;
                long value = nonzero ? vlc.ReadAbsolute(bits) - 1L : 0;
                int refinement = _model.Bits[i == 0 ? 0 : 1];
                value = (value << refinement) | bits.Read(refinement);
                if (value > int.MaxValue) throw new FormatException("JPEG-XR DC coefficient exceeds the managed range.");
                destination[offset + i] = value == 0 ? 0 : bits.Flag() ? -(int)value : (int)value;
            }
            _model.Update(lumaCount, chromaCount);
            if (adapt) { _luma.Adapt(); _chroma.Adapt(); }
        }
    }
}
