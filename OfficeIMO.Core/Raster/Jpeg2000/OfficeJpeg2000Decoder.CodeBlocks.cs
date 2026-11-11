namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    /// <summary>
    /// Embedded block decoder (Annex D) with the MQ arithmetic decoder (Annex C). One instance is reused for every
    /// code-block of an image; its scratch buffers are sized for the largest permitted code-block.
    /// </summary>
    private sealed class CodeBlockDecoder {
        private const int Significant = 1, Negative = 2, Visited = 4, Refined = 8;
        private const int ContextRunLength = 17, ContextUniform = 18;
        private const int MaximumCoefficients = 4096;
        private const int MaximumFlags = 1026 * 6;

        private static readonly byte[] ZeroCodingLabels = BuildZeroCodingLabels();
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

        private readonly int[] _magnitudes = new int[MaximumCoefficients];
        private readonly byte[] _flags = new byte[MaximumFlags];
        private readonly byte[] _states = new byte[19];
        private readonly byte[] _mps = new byte[19];
        private byte[] _buffer = new byte[1024];
        private readonly Budget _budget;

        private int _width, _height, _stride, _group;
        private bool _causal;
        private int _lastPlane, _lastKind;

        // MQ decoder registers (C.3).
        private byte[] _data = Array.Empty<byte>();
        private int _position, _end;
        private uint _c, _a;
        private int _ct;
        private int _rawCurrent, _rawBits;

        internal CodeBlockDecoder(Budget budget) {
            _budget = budget;
            budget.Reserve(MaximumCoefficients * 4L + MaximumFlags + _buffer.Length);
        }

        private static byte[] BuildZeroCodingLabels() {
            var labels = new byte[3 * 45];
            for (int h = 0; h <= 2; h++) {
                for (int v = 0; v <= 2; v++) {
                    for (int d = 0; d <= 4; d++) {
                        labels[h * 15 + v * 5 + d] = HorizontalLabel(h, v, d);
                        labels[45 + h * 15 + v * 5 + d] = HorizontalLabel(v, h, d);
                        int hv = h + v;
                        labels[90 + h * 15 + v * 5 + d] = (byte)(d >= 3 ? 8 : d == 2 ? (hv >= 1 ? 7 : 6) :
                            d == 1 ? (hv >= 2 ? 5 : hv == 1 ? 4 : 3) : (hv >= 2 ? 2 : hv == 1 ? 1 : 0));
                    }
                }
            }
            return labels;

            // Table D.1 for LL and LH subbands; HL swaps the horizontal and vertical sums.
            static byte HorizontalLabel(int h, int v, int d) => (byte)(h == 2 ? 8 :
                h == 1 ? (v >= 1 ? 7 : d >= 1 ? 6 : 5) :
                v == 2 ? 4 : v == 1 ? 3 : d >= 2 ? 2 : d == 1 ? 1 : 0);
        }

        /// <summary>Decodes every included pass of the code-block into the magnitude and sign scratch buffers.</summary>
        internal bool Decode(CodeBlock block, SubBand band, int style, byte[] source) {
            _width = block.X1 - block.X0;
            _height = block.Y1 - block.Y0;
            _stride = _width + 2;
            int planes = band.MagnitudeBits - block.ZeroPlanes;
            if (block.Passes == 0 || planes <= 0 || block.Segments == null || _width <= 0 || _height <= 0) return false;
            Array.Clear(_magnitudes, 0, _width * _height);
            Array.Clear(_flags, 0, _stride * (_height + 2));
            _group = band.Orientation == 1 ? 1 : band.Orientation == 3 ? 2 : 0;
            _causal = (style & CodeBlockVerticallyCausal) != 0;
            bool bypass = (style & CodeBlockBypass) != 0;
            ResetContexts();
            int plane = planes - 1, kind = 2, passIndex = 0;
            _lastPlane = plane;
            _lastKind = 2;
            long area = (long)_width * _height;
            foreach (CodeSegment segment in block.Segments) {
                LoadSegment(segment, source);
                bool raw = bypass && passIndex >= 10 && kind != 2;
                if (raw) {
                    _rawCurrent = 0;
                    _rawBits = 0;
                } else {
                    InitializeArithmetic();
                }
                for (int pass = 0; pass < segment.Passes; pass++) {
                    if (plane < 0) return true;
                    _budget.Charge(area);
                    if (kind == 0) SignificancePass(plane, raw);
                    else if (kind == 1) RefinementPass(plane, raw);
                    else {
                        CleanupPass(plane);
                        if ((style & CodeBlockSegmentationSymbols) != 0) {
                            for (int index = 0; index < 4; index++) DecodeDecision(ContextUniform);
                        }
                    }
                    _lastPlane = plane;
                    _lastKind = kind;
                    if (kind == 2) {
                        plane--;
                        kind = 0;
                    } else {
                        kind++;
                    }
                    if ((style & CodeBlockReset) != 0) ResetContexts();
                    passIndex++;
                }
            }
            return true;
        }

        private void LoadSegment(CodeSegment segment, byte[] source) {
            if (segment.Length > _buffer.Length) {
                int size = Math.Max(segment.Length, _buffer.Length * 2);
                _budget.Reserve(size);
                _budget.Release(_buffer.Length);
                _buffer = new byte[size];
            }
            int offset = 0;
            foreach ((int start, int length) in segment.Chunks) {
                Buffer.BlockCopy(source, start, _buffer, offset, length);
                offset += length;
            }
            _data = _buffer;
            _position = 0;
            _end = offset;
        }

        private void ResetContexts() {
            Array.Clear(_states, 0, _states.Length);
            Array.Clear(_mps, 0, _mps.Length);
            _states[0] = 4;
            _states[ContextRunLength] = 3;
            _states[ContextUniform] = 46;
        }

        private int ByteAt(int position) => position < _end ? _data[position] : 0xFF;

        // INITDEC (C.3.5).
        private void InitializeArithmetic() {
            _c = (uint)ByteAt(_position) << 16;
            ByteIn();
            _c <<= 7;
            _ct -= 7;
            _a = 0x8000;
        }

        // BYTEIN (C.3.4).
        private void ByteIn() {
            if (ByteAt(_position) == 0xFF) {
                if (ByteAt(_position + 1) > 0x8F) {
                    _c += 0xFF00;
                    _ct = 8;
                } else {
                    _position++;
                    _c += (uint)ByteAt(_position) << 9;
                    _ct = 7;
                }
            } else {
                _position++;
                _c += (uint)ByteAt(_position) << 8;
                _ct = 8;
            }
        }

        // DECODE (C.3.2) with conditional exchange; the LPS sub-interval lies at the bottom.
        private int DecodeDecision(int context) {
            int state = _states[context];
            int mps = _mps[context];
            uint qe = Qe[state];
            int decision;
            _a -= qe;
            if ((_c >> 16) < qe) {
                if (_a < qe) {
                    decision = mps;
                    _states[context] = NextMps[state];
                } else {
                    decision = 1 - mps;
                    if (Switch[state] != 0) _mps[context] = (byte)(1 - mps);
                    _states[context] = NextLps[state];
                }
                _a = qe;
            } else {
                _c -= qe << 16;
                if ((_a & 0x8000) != 0) return mps;
                if (_a < qe) {
                    decision = 1 - mps;
                    if (Switch[state] != 0) _mps[context] = (byte)(1 - mps);
                    _states[context] = NextLps[state];
                } else {
                    decision = mps;
                    _states[context] = NextMps[state];
                }
            }
            do {
                if (_ct == 0) ByteIn();
                _a <<= 1;
                _c <<= 1;
                _ct--;
            } while ((_a & 0x8000) == 0);
            return decision;
        }

        // Raw (bypass) bits (D.6) with a stuffed zero bit after 0xFF.
        private int RawBit() {
            if (_rawBits == 0) {
                int value = _position < _end ? _data[_position++] : 0xFF;
                _rawBits = _rawCurrent == 0xFF ? 7 : 8;
                _rawCurrent = value;
            }
            _rawBits--;
            return (_rawCurrent >> _rawBits) & 1;
        }

        private int ZeroCodingLabel(int index, bool lastStripeRow) {
            byte[] f = _flags;
            int s = _stride;
            bool cut = lastStripeRow && _causal;
            int h = (f[index - 1] & Significant) + (f[index + 1] & Significant);
            int v = (f[index - s] & Significant) + (cut ? 0 : f[index + s] & Significant);
            int d = (f[index - s - 1] & Significant) + (f[index - s + 1] & Significant) +
                (cut ? 0 : (f[index + s - 1] & Significant) + (f[index + s + 1] & Significant));
            return ZeroCodingLabels[_group * 45 + h * 15 + v * 5 + d];
        }

        private static int SignContribution(byte flags) => (flags & Significant) == 0 ? 0 : (flags & Negative) != 0 ? -1 : 1;

        // Sign coding (D.3.2, Tables D.2 and D.3).
        private int DecodeSign(int index, bool lastStripeRow, bool raw) {
            if (raw) return RawBit();
            byte[] f = _flags;
            int h = SignContribution(f[index - 1]) + SignContribution(f[index + 1]);
            int v = SignContribution(f[index - _stride]) + (lastStripeRow && _causal ? 0 : SignContribution(f[index + _stride]));
            h = h < -1 ? -1 : h > 1 ? 1 : h;
            v = v < -1 ? -1 : v > 1 ? 1 : v;
            int context, inversion = 0;
            if (h == 0) {
                context = v == 0 ? 9 : 10;
                if (v < 0) inversion = 1;
            } else {
                context = v == 0 ? 12 : v == h ? 13 : 11;
                if (h < 0) inversion = 1;
            }
            return DecodeDecision(context) ^ inversion;
        }

        private bool HasSignificantNeighbour(int index, bool lastStripeRow) {
            byte[] f = _flags;
            int s = _stride;
            int any = f[index - 1] | f[index + 1] | f[index - s] | f[index - s - 1] | f[index - s + 1];
            if (!(lastStripeRow && _causal)) any |= f[index + s] | f[index + s - 1] | f[index + s + 1];
            return (any & Significant) != 0;
        }

        private void SignificancePass(int plane, bool raw) {
            int one = 1 << plane;
            for (int stripe = 0; stripe < _height; stripe += 4) {
                int rows = Math.Min(4, _height - stripe);
                for (int x = 0; x < _width; x++) {
                    for (int row = 0; row < rows; row++) {
                        int y = stripe + row;
                        int index = (y + 1) * _stride + x + 1;
                        if ((_flags[index] & Significant) != 0) continue;
                        bool last = row == 3;
                        int label = ZeroCodingLabel(index, last);
                        if (label == 0) continue;
                        _flags[index] |= Visited;
                        int bit = raw ? RawBit() : DecodeDecision(label);
                        if (bit == 0) continue;
                        int sign = DecodeSign(index, last, raw);
                        _magnitudes[y * _width + x] = one;
                        _flags[index] |= (byte)(Significant | (sign != 0 ? Negative : 0));
                    }
                }
            }
        }

        private void RefinementPass(int plane, bool raw) {
            int one = 1 << plane;
            for (int stripe = 0; stripe < _height; stripe += 4) {
                int rows = Math.Min(4, _height - stripe);
                for (int x = 0; x < _width; x++) {
                    for (int row = 0; row < rows; row++) {
                        int y = stripe + row;
                        int index = (y + 1) * _stride + x + 1;
                        byte flags = _flags[index];
                        if ((flags & (Significant | Visited)) != Significant) continue;
                        int bit;
                        if (raw) {
                            bit = RawBit();
                        } else {
                            int context = (flags & Refined) != 0 ? 16 : HasSignificantNeighbour(index, row == 3) ? 15 : 14;
                            bit = DecodeDecision(context);
                        }
                        if (bit != 0) _magnitudes[y * _width + x] |= one;
                        _flags[index] = (byte)(flags | Refined);
                    }
                }
            }
        }

        private void CleanupPass(int plane) {
            int one = 1 << plane;
            for (int stripe = 0; stripe < _height; stripe += 4) {
                int rows = Math.Min(4, _height - stripe);
                for (int x = 0; x < _width; x++) {
                    int first = (stripe + 1) * _stride + x + 1;
                    int start = 0;
                    if (rows == 4 && IsRunCandidate(first)) {
                        if (DecodeDecision(ContextRunLength) == 0) {
                            start = 4;
                        } else {
                            int position = DecodeDecision(ContextUniform) << 1;
                            position |= DecodeDecision(ContextUniform);
                            int index = first + position * _stride;
                            int sign = DecodeSign(index, position == 3, raw: false);
                            _magnitudes[(stripe + position) * _width + x] = one;
                            _flags[index] |= (byte)(Significant | (sign != 0 ? Negative : 0));
                            start = position + 1;
                        }
                    }
                    for (int row = start; row < rows; row++) {
                        int index = first + row * _stride;
                        if ((_flags[index] & (Significant | Visited)) != 0) continue;
                        bool last = row == 3;
                        if (DecodeDecision(ZeroCodingLabel(index, last)) == 0) continue;
                        int sign = DecodeSign(index, last, raw: false);
                        _magnitudes[(stripe + row) * _width + x] = one;
                        _flags[index] |= (byte)(Significant | (sign != 0 ? Negative : 0));
                    }
                    for (int row = 0; row < rows; row++) _flags[first + row * _stride] &= unchecked((byte)~Visited);
                }
            }
        }

        private bool IsRunCandidate(int first) {
            for (int row = 0; row < 4; row++) {
                int index = first + row * _stride;
                if ((_flags[index] & (Significant | Visited)) != 0 || ZeroCodingLabel(index, row == 3) != 0) return false;
            }
            return true;
        }

        // Lowest decoded bit-plane of a coefficient, used for mid-point reconstruction (E.1.1.2).
        private int LowestPlane(int index) {
            if (_lastKind != 0) return _lastPlane;
            return (_flags[index] & Visited) != 0 ? _lastPlane : _lastPlane + 1;
        }

        internal void WriteIntegers(int[] target, int stride, int left, int top, int regionShift) {
            for (int y = 0; y < _height; y++) {
                int row = (top + y) * stride + left;
                for (int x = 0; x < _width; x++) {
                    int magnitude = _magnitudes[y * _width + x];
                    if (magnitude == 0) continue;
                    int index = (y + 1) * _stride + x + 1;
                    int plane = LowestPlane(index);
                    if (regionShift > 0 && magnitude >= 1 << regionShift) {
                        magnitude >>= regionShift;
                        plane = Math.Max(0, plane - regionShift);
                    }
                    if (plane > 0) magnitude += 1 << (plane - 1);
                    target[row + x] = (_flags[index] & Negative) != 0 ? -magnitude : magnitude;
                }
            }
        }

        internal void WriteFloats(float[] target, int stride, int left, int top, int regionShift, float step) {
            for (int y = 0; y < _height; y++) {
                int row = (top + y) * stride + left;
                for (int x = 0; x < _width; x++) {
                    int magnitude = _magnitudes[y * _width + x];
                    if (magnitude == 0) continue;
                    int index = (y + 1) * _stride + x + 1;
                    int plane = LowestPlane(index);
                    if (regionShift > 0 && magnitude >= 1 << regionShift) {
                        magnitude >>= regionShift;
                        plane = Math.Max(0, plane - regionShift);
                    }
                    float value = (magnitude + (1 << plane) * 0.5F) * step;
                    target[row + x] = (_flags[index] & Negative) != 0 ? -value : value;
                }
            }
        }
    }
}
