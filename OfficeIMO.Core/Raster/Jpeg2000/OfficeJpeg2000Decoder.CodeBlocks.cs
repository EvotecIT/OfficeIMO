namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    /// <summary>
    /// Embedded block decoder (Annex D) with the MQ arithmetic decoder (Annex C). One instance is reused for every
    /// code-block of an image; its scratch buffers are sized for the largest permitted code-block.
    /// </summary>
    private sealed class CodeBlockDecoder {
        // Coefficient state: its own flags, then the significance of its eight neighbours and the signs of the four
        // direct ones, updated whenever a coefficient becomes significant so that contexts are single lookups.
        private const int Significant = 1, Negative = 2, Visited = 4, Refined = 8;
        private const int NorthSignificant = 1 << 4, SouthSignificant = 1 << 5, WestSignificant = 1 << 6, EastSignificant = 1 << 7,
            NorthWestSignificant = 1 << 8, NorthEastSignificant = 1 << 9, SouthWestSignificant = 1 << 10, SouthEastSignificant = 1 << 11,
            NorthNegative = 1 << 12, SouthNegative = 1 << 13, WestNegative = 1 << 14, EastNegative = 1 << 15;
        private const int Neighbours = 0x0FF0;
        // Vertically causal mode ignores the next stripe for the last row of a stripe (D.7).
        private const int SouthNeighbours = SouthSignificant | SouthWestSignificant | SouthEastSignificant | SouthNegative;
        private const int ContextRunLength = 17, ContextUniform = 18;
        private const int MaximumCoefficients = 4096;
        private const int MaximumFlags = 1026 * 6;

        private static readonly byte[] ZeroCodingLabels = BuildZeroCodingLabels();
        private static readonly byte[] SignContexts = BuildSignContexts();
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
        private readonly ushort[] _flags = new ushort[MaximumFlags];
        private readonly byte[] _contexts = new byte[19]; // MQ state index times two plus the MPS
        private byte[] _buffer = new byte[1024];
        private readonly Budget _budget;

        private int _width, _height, _stride, _labelOffset;
        private int _lastRowMask; // flags mask for the last row of a stripe
        private bool _causal;
        private int _lastPlane, _lastKind;

        // MQ decoder registers (C.3).
        private byte[] _data = Array.Empty<byte>();
        private int _end;
        private MqState _mq;
        private int _rawCurrent, _rawBits;

        /// <summary>MQ decoder registers (C.3) and the read position, copied to a local by the coding passes.</summary>
        private struct MqState {
            internal uint A, C;
            internal int Ct, Position;
        }

        internal CodeBlockDecoder(Budget budget) {
            _budget = budget;
            budget.Reserve(MaximumCoefficients * 4L + MaximumFlags * 2L + _buffer.Length);
        }

        // Table D.1 by the eight neighbour-significance bits: LL and LH, then HL (horizontal and vertical swapped), then HH.
        private static byte[] BuildZeroCodingLabels() {
            var labels = new byte[3 * 256];
            for (int bits = 0; bits < 256; bits++) {
                int v = (bits & 1) + ((bits >> 1) & 1);
                int h = ((bits >> 2) & 1) + ((bits >> 3) & 1);
                int d = ((bits >> 4) & 1) + ((bits >> 5) & 1) + ((bits >> 6) & 1) + ((bits >> 7) & 1);
                int hv = h + v;
                labels[bits] = HorizontalLabel(h, v, d);
                labels[256 + bits] = HorizontalLabel(v, h, d);
                labels[512 + bits] = (byte)(d >= 3 ? 8 : d == 2 ? (hv >= 1 ? 7 : 6) :
                    d == 1 ? (hv >= 2 ? 5 : hv == 1 ? 4 : 3) : (hv >= 2 ? 2 : hv == 1 ? 1 : 0));
            }
            return labels;

            static byte HorizontalLabel(int h, int v, int d) => (byte)(h == 2 ? 8 :
                h == 1 ? (v >= 1 ? 7 : d >= 1 ? 6 : 5) :
                v == 2 ? 4 : v == 1 ? 3 : d >= 2 ? 2 : d == 1 ? 1 : 0);
        }

        // Tables D.2 and D.3 by the significance (low nibble) and sign (high nibble) of the N, S, W and E neighbours:
        // the context, with the XOR bit in bit 7.
        private static byte[] BuildSignContexts() {
            var contexts = new byte[256];
            for (int bits = 0; bits < 256; bits++) {
                int v = Contribution(bits, 0) + Contribution(bits, 1);
                int h = Contribution(bits, 2) + Contribution(bits, 3);
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
                contexts[bits] = (byte)(context | inversion << 7);
            }
            return contexts;

            static int Contribution(int bits, int neighbour) => ((bits >> neighbour) & 1) == 0 ? 0 : ((bits >> (neighbour + 4)) & 1) != 0 ? -1 : 1;
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
            _labelOffset = band.Orientation == 1 ? 256 : band.Orientation == 3 ? 512 : 0;
            _causal = (style & CodeBlockVerticallyCausal) != 0;
            _lastRowMask = _causal ? ~SouthNeighbours : ~0;
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
                            MqState mq = _mq;
                            for (int index = 0; index < 4; index++) Decide(ref mq, ContextUniform);
                            _mq = mq;
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
            _mq.Position = 0;
            _end = offset;
        }

        private void ResetContexts() {
            Array.Clear(_contexts, 0, _contexts.Length);
            _contexts[0] = 4 << 1;
            _contexts[ContextRunLength] = 3 << 1;
            _contexts[ContextUniform] = 46 << 1;
        }

        private int ByteAt(int position) => position < _end ? _data[position] : 0xFF;

        // INITDEC (C.3.5).
        private void InitializeArithmetic() {
            MqState mq = _mq;
            mq.C = (uint)ByteAt(mq.Position) << 16;
            ByteIn(ref mq);
            mq.C <<= 7;
            mq.Ct -= 7;
            mq.A = 0x8000;
            _mq = mq;
        }

        // BYTEIN (C.3.4).
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        private void ByteIn(ref MqState mq) {
            if (ByteAt(mq.Position) == 0xFF) {
                if (ByteAt(mq.Position + 1) > 0x8F) {
                    mq.C += 0xFF00;
                    mq.Ct = 8;
                } else {
                    mq.Position++;
                    mq.C += (uint)ByteAt(mq.Position) << 9;
                    mq.Ct = 7;
                }
            } else {
                mq.Position++;
                mq.C += (uint)ByteAt(mq.Position) << 8;
                mq.Ct = 8;
            }
        }

        // Qe and the next context index after an MPS or LPS decision, by state index times two plus the MPS (Table C.2).
        private static readonly ushort[] IndexQe = BuildIndexTable(static (state, mps) => Qe[state]);
        private static readonly byte[] IndexAfterMps = Array.ConvertAll(BuildIndexTable(static (state, mps) => (ushort)(NextMps[state] << 1 | mps)), static value => (byte)value);
        private static readonly byte[] IndexAfterLps = Array.ConvertAll(BuildIndexTable(static (state, mps) => (ushort)(NextLps[state] << 1 | (mps ^ Switch[state]))), static value => (byte)value);

        private static ushort[] BuildIndexTable(Func<int, int, ushort> value) {
            var table = new ushort[Qe.Length * 2];
            for (int index = 0; index < table.Length; index++) table[index] = value(index >> 1, index & 1);
            return table;
        }

        // DECODE (C.3.2) with conditional exchange; the LPS sub-interval lies at the bottom.
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        private int Decide(ref MqState mq, int context) {
            byte[] contexts = _contexts;
            int index = contexts[context];
            uint qe = IndexQe[index];
            uint a = mq.A - qe;
            int decision;
            if ((mq.C >> 16) < qe) {
                if (a < qe) {
                    decision = index & 1;
                    contexts[context] = IndexAfterMps[index];
                } else {
                    decision = (index & 1) ^ 1;
                    contexts[context] = IndexAfterLps[index];
                }
                a = qe;
            } else {
                mq.C -= qe << 16;
                if ((a & 0x8000) != 0) {
                    mq.A = a;
                    return index & 1;
                }
                if (a < qe) {
                    decision = (index & 1) ^ 1;
                    contexts[context] = IndexAfterLps[index];
                } else {
                    decision = index & 1;
                    contexts[context] = IndexAfterMps[index];
                }
            }
            do {
                if (mq.Ct == 0) ByteIn(ref mq);
                a <<= 1;
                mq.C <<= 1;
                mq.Ct--;
            } while ((a & 0x8000) == 0);
            mq.A = a;
            return decision;
        }

        // Raw (bypass) bits (D.6) with a stuffed zero bit after 0xFF.
        private int RawBit() {
            if (_rawBits == 0) {
                int value = _mq.Position < _end ? _data[_mq.Position++] : 0xFF;
                _rawBits = _rawCurrent == 0xFF ? 7 : 8;
                _rawCurrent = value;
            }
            _rawBits--;
            return (_rawCurrent >> _rawBits) & 1;
        }

        // Sign coding (D.3.2) from the neighbour bits of a coefficient.
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        private int DecodeSign(ref MqState mq, int flags) {
            int entry = SignContexts[((flags >> 4) & 0x0F) | ((flags >> 8) & 0xF0)];
            return Decide(ref mq, entry & 0x1F) ^ (entry >> 7);
        }

        private void MakeSignificant(int index, int sign) {
            ushort[] f = _flags;
            int s = _stride;
            if (sign != 0) {
                f[index] |= Significant | Negative;
                f[index - s] |= SouthSignificant | SouthNegative;
                f[index + s] |= NorthSignificant | NorthNegative;
                f[index - 1] |= EastSignificant | EastNegative;
                f[index + 1] |= WestSignificant | WestNegative;
            } else {
                f[index] |= Significant;
                f[index - s] |= SouthSignificant;
                f[index + s] |= NorthSignificant;
                f[index - 1] |= EastSignificant;
                f[index + 1] |= WestSignificant;
            }
            f[index - s - 1] |= SouthEastSignificant;
            f[index - s + 1] |= SouthWestSignificant;
            f[index + s - 1] |= NorthEastSignificant;
            f[index + s + 1] |= NorthWestSignificant;
        }

        private void SignificancePass(int plane, bool raw) {
            int one = 1 << plane;
            ushort[] f = _flags;
            int stride = _stride;
            // Raw passes read through the field; only arithmetic passes work on, and write back, the local copy.
            MqState mq = _mq;
            for (int stripe = 0; stripe < _height; stripe += 4) {
                int rows = Math.Min(4, _height - stripe);
                int first = (stripe + 1) * stride + 1;
                for (int x = 0; x < _width; x++) {
                    int top = first + x;
                    // A column without significant neighbours, or with only significant coefficients, codes nothing here.
                    if (rows == 4) {
                        int f0 = f[top], f1 = f[top + stride], f2 = f[top + 2 * stride], f3 = f[top + 3 * stride];
                        if (((f0 | f1 | f2 | f3) & Neighbours) == 0 || (f0 & f1 & f2 & f3 & Significant) != 0) continue;
                    }
                    for (int row = 0, index = top; row < rows; row++, index += stride) {
                        int flags = f[index];
                        if (row == 3) flags &= _lastRowMask;
                        if ((flags & Significant) != 0 || (flags & Neighbours) == 0) continue;
                        f[index] |= Visited;
                        int bit = raw ? RawBit() : Decide(ref mq, ZeroCodingLabels[_labelOffset + ((flags >> 4) & 0xFF)]);
                        if (bit == 0) continue;
                        int sign = raw ? RawBit() : DecodeSign(ref mq, flags);
                        _magnitudes[(stripe + row) * _width + x] = one;
                        MakeSignificant(index, sign);
                    }
                }
            }
            if (!raw) _mq = mq;
        }

        private void RefinementPass(int plane, bool raw) {
            int one = 1 << plane;
            ushort[] f = _flags;
            int stride = _stride;
            MqState mq = _mq;
            for (int stripe = 0; stripe < _height; stripe += 4) {
                int rows = Math.Min(4, _height - stripe);
                int first = (stripe + 1) * stride + 1;
                for (int x = 0; x < _width; x++) {
                    int top = first + x;
                    if (rows == 4 && ((f[top] | f[top + stride] | f[top + 2 * stride] | f[top + 3 * stride]) & Significant) == 0) continue;
                    for (int row = 0, index = top; row < rows; row++, index += stride) {
                        int flags = f[index];
                        if ((flags & (Significant | Visited)) != Significant) continue;
                        int bit;
                        if (raw) {
                            bit = RawBit();
                        } else {
                            int neighbours = row == 3 ? flags & _lastRowMask & Neighbours : flags & Neighbours;
                            bit = Decide(ref mq, (flags & Refined) != 0 ? 16 : neighbours != 0 ? 15 : 14);
                        }
                        if (bit != 0) _magnitudes[(stripe + row) * _width + x] |= one;
                        f[index] = (ushort)(flags | Refined);
                    }
                }
            }
            if (!raw) _mq = mq;
        }

        private void CleanupPass(int plane) {
            int one = 1 << plane;
            ushort[] f = _flags;
            int stride = _stride;
            MqState mq = _mq;
            for (int stripe = 0; stripe < _height; stripe += 4) {
                int rows = Math.Min(4, _height - stripe);
                for (int x = 0; x < _width; x++) {
                    int first = (stripe + 1) * stride + x + 1;
                    int start = 0;
                    if (rows == 4) {
                        int f0 = f[first], f1 = f[first + stride], f2 = f[first + 2 * stride], f3 = f[first + 3 * stride];
                        // Every coefficient already significant or coded in this bit-plane: only the visited marks change.
                        if (((f0 | f0 >> 2) & (f1 | f1 >> 2) & (f2 | f2 >> 2) & (f3 | f3 >> 2) & Significant) != 0) {
                            f[first] = (ushort)(f0 & ~Visited);
                            f[first + stride] = (ushort)(f1 & ~Visited);
                            f[first + 2 * stride] = (ushort)(f2 & ~Visited);
                            f[first + 3 * stride] = (ushort)(f3 & ~Visited);
                            continue;
                        }
                    }
                    if (rows == 4 &&
                        ((f[first] | f[first + stride] | f[first + 2 * stride] | (f[first + 3 * stride] & _lastRowMask)) &
                         (Significant | Visited | Neighbours)) == 0) {
                        if (Decide(ref mq, ContextRunLength) == 0) {
                            start = 4;
                        } else {
                            int position = Decide(ref mq, ContextUniform) << 1;
                            position |= Decide(ref mq, ContextUniform);
                            int index = first + position * stride;
                            int flags = f[index];
                            if (position == 3) flags &= _lastRowMask;
                            int sign = DecodeSign(ref mq, flags);
                            _magnitudes[(stripe + position) * _width + x] = one;
                            MakeSignificant(index, sign);
                            start = position + 1;
                        }
                    }
                    for (int row = start, index = first + start * stride; row < rows; row++, index += stride) {
                        int flags = f[index];
                        if ((flags & (Significant | Visited)) != 0) continue;
                        if (row == 3) flags &= _lastRowMask;
                        if (Decide(ref mq, ZeroCodingLabels[_labelOffset + ((flags >> 4) & 0xFF)]) == 0) continue;
                        int sign = DecodeSign(ref mq, flags);
                        _magnitudes[(stripe + row) * _width + x] = one;
                        MakeSignificant(index, sign);
                    }
                    for (int row = 0, index = first; row < rows; row++, index += stride) f[index] &= unchecked((ushort)~Visited);
                }
            }
            _mq = mq;
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
