#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.Arm;
using System.Runtime.Intrinsics.X86;
#endif

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    private const float IrreversibleAlpha = -1.586134342059924F;
    private const float IrreversibleBeta = -0.052980118572961F;
    private const float IrreversibleGamma = 0.882911075530934F;
    private const float IrreversibleDelta = 0.443506852043971F;
    private const float IrreversibleK = 1.230174104914001F;

    /// <summary>Rows a resolution level synthesizes at least per request, to amortize the filter margins.</summary>
    private const int MinimumLevelRows = 16;

    /// <summary>State shared by every tile of one decode.</summary>
    private sealed class Workspace {
        internal Workspace(Budget budget, CodeBlockDecoder blocks) {
            Budget = budget;
            Blocks = blocks;
            Pool = new SamplePool(budget);
        }

        internal Budget Budget { get; }
        internal CodeBlockDecoder Blocks { get; }
        internal SamplePool Pool { get; }
        internal byte[] Source = Array.Empty<byte>();
    }

    /// <summary>Row buffers reused across subbands, levels and tiles; each new array is charged to the budget once.</summary>
    private sealed class SamplePool {
        private readonly Budget _budget;
        private readonly List<int[]> _integers = new();
        private readonly List<float[]> _floats = new();

        internal SamplePool(Budget budget) => _budget = budget;

        internal int[] RentIntegers(int length) {
            int index = Smallest(_integers, length);
            if (index >= 0) return Take(_integers, index);
            _budget.Reserve(length * 4L);
            return new int[length];
        }

        internal float[] RentFloats(int length) {
            int index = Smallest(_floats, length);
            if (index >= 0) return Take(_floats, index);
            _budget.Reserve(length * 4L);
            return new float[length];
        }

        internal void Return(int[]? array) {
            if (array != null && array.Length > 0) _integers.Add(array);
        }

        internal void Return(float[]? array) {
            if (array != null && array.Length > 0) _floats.Add(array);
        }

        private static int Smallest<T>(List<T[]> list, int length) {
            int best = -1;
            for (int index = 0; index < list.Count; index++) {
                if (list[index].Length >= length && (best < 0 || list[index].Length < list[best].Length)) best = index;
            }
            return best;
        }

        private static T[] Take<T>(List<T[]> list, int index) {
            T[] array = list[index];
            list[index] = list[list.Count - 1];
            list.RemoveAt(list.Count - 1);
            return array;
        }
    }

    /// <summary>
    /// A window of consecutive rows of one subband or resolution level, produced top to bottom on request. Rows
    /// before the requested range are dropped, so a source holds only what its consumer still needs.
    /// </summary>
    private abstract class RowSource {
        protected RowSource(int width, int height, bool reversible) {
            Width = Math.Max(0, width);
            Height = Math.Max(0, height);
            Reversible = reversible;
        }

        internal int Width { get; }
        internal int Height { get; }
        internal bool Reversible { get; }
        internal int First { get; private set; }
        internal int End { get; private set; }
        internal int[]? Ints;
        internal float[]? Floats;
        private int _base; // buffer row holding First; rows are moved to the front only when the buffer is full

        /// <summary>Buffer row of a held row.</summary>
        internal int RowIndex(int row) => row - First + _base;

        internal int Offset(int row) => (row - First + _base) * Width;

        /// <summary>Makes rows <paramref name="first"/> to <paramref name="last"/> available.</summary>
        internal void Ensure(int first, int last, Workspace workspace) {
            if (first < First || first > last || last >= Height) throw new Jpeg2000Exception("The JPEG 2000 wavelet rows were requested out of order.");
            int drop = Math.Min(first, End) - First;
            if (drop > 0) {
                First += drop;
                _base = First == End ? 0 : _base + drop;
            }
            if (Width == 0) {
                // Empty rows carry no samples and no code-blocks.
                End = Math.Max(End, last + 1);
                return;
            }
            while (End <= last) {
                workspace.Budget.Token.ThrowIfCancellationRequested();
                int end = Produce(last, workspace);
                if (end <= End) throw new Jpeg2000Exception("The JPEG 2000 wavelet rows could not be produced.");
                End = end;
            }
        }

        /// <summary>Appends rows from <see cref="End"/> onwards, covering at least <paramref name="last"/>; returns the new end.</summary>
        protected abstract int Produce(int last, Workspace workspace);

        /// <summary>Makes room for rows from <see cref="First"/> up to <paramref name="end"/>.</summary>
        protected void Reserve(int end, SamplePool pool) {
            long needed = (long)(end - First) * Width;
            if (needed > int.MaxValue) throw new Jpeg2000Exception("The JPEG 2000 tile-component is too wide.");
            long used = (long)(end - First + _base) * Width;
            int start = _base * Width, held = (End - First) * Width;
            if (Reversible) {
                int length = Ints?.Length ?? 0;
                if (used <= length) return;
                if (needed <= length) {
                    Array.Copy(Ints!, start, Ints!, 0, held);
                } else {
                    int[] next = pool.RentIntegers((int)Math.Max(needed, Math.Min(int.MaxValue, needed * 3 / 2)));
                    if (Ints != null) {
                        Array.Copy(Ints, start, next, 0, held);
                        pool.Return(Ints);
                    }
                    Ints = next;
                }
            } else {
                int length = Floats?.Length ?? 0;
                if (used <= length) return;
                if (needed <= length) {
                    Array.Copy(Floats!, start, Floats!, 0, held);
                } else {
                    float[] next = pool.RentFloats((int)Math.Max(needed, Math.Min(int.MaxValue, needed * 3 / 2)));
                    if (Floats != null) {
                        Array.Copy(Floats, start, next, 0, held);
                        pool.Return(Floats);
                    }
                    Floats = next;
                }
            }
            _base = 0;
        }

        internal virtual void Return(SamplePool pool) {
            pool.Return(Ints);
            pool.Return(Floats);
            Ints = null;
            Floats = null;
            _base = 0;
        }
    }

    private static RowSource CreateComponentSource(TileComponent component, Workspace workspace) {
        Resolution[] resolutions = component.Resolutions;
        bool reversible = component.Coding.Reversible;
        RowSource source = new BandSource(component, resolutions[0], 0, workspace);
        for (int r = 1; r < resolutions.Length; r++) {
            Resolution resolution = resolutions[r];
            if (resolution.Width > 0 && resolution.Height > 0) workspace.Budget.Charge((long)resolution.Width * resolution.Height * 2);
            source = new LevelSource(source, new BandSource(component, resolution, 0, workspace), new BandSource(component, resolution, 1, workspace),
                new BandSource(component, resolution, 2, workspace), resolution, reversible);
        }
        return source;
    }

    /// <summary>Dequantized coefficients of one subband, decoded one row of code-blocks at a time.</summary>
    private sealed class BandSource : RowSource {
        private readonly SubBand _band;
        private readonly int _style;
        private readonly int _regionShift;
        private readonly CodeBlock[] _blocks;
        private readonly int[] _rowStarts; // _blocks index of each code-block row, plus one end entry
        private readonly int _blockRowShift;
        private readonly long _firstBlockRow;
        private int _nextBlockRow;

        internal BandSource(TileComponent component, Resolution resolution, int bandIndex, Workspace workspace)
            : base(resolution.Bands[bandIndex].X1 - resolution.Bands[bandIndex].X0, resolution.Bands[bandIndex].Y1 - resolution.Bands[bandIndex].Y0,
                component.Coding.Reversible) {
            _band = resolution.Bands[bandIndex];
            _style = component.Coding.Style;
            _regionShift = component.RegionShift;
            _blockRowShift = _band.CodeBlockHeight;
            _firstBlockRow = _band.Y0 >> _blockRowShift;
            int rows = Height > 0 && Width > 0 ? (int)(CeilDiv(_band.Y1, 1L << _blockRowShift) - _firstBlockRow) : 0;
            _rowStarts = new int[rows + 1];
            int total = 0;
            foreach (Precinct precinct in resolution.Precincts) {
                foreach (CodeBlock block in precinct.Bands[bandIndex].Blocks) {
                    int row = (int)((block.Y0 >> _blockRowShift) - _firstBlockRow);
                    if ((uint)row >= (uint)rows) throw new Jpeg2000Exception("A JPEG 2000 code-block lies outside its subband.");
                    _rowStarts[row + 1]++;
                    total++;
                }
            }
            for (int row = 0; row < rows; row++) _rowStarts[row + 1] += _rowStarts[row];
            _blocks = total == 0 ? Array.Empty<CodeBlock>() : new CodeBlock[total];
            if (total == 0) return;
            var next = new int[rows];
            Array.Copy(_rowStarts, next, rows);
            foreach (Precinct precinct in resolution.Precincts) {
                foreach (CodeBlock block in precinct.Bands[bandIndex].Blocks) {
                    int row = (int)((block.Y0 >> _blockRowShift) - _firstBlockRow);
                    _blocks[next[row]++] = block;
                }
            }
        }

        protected override int Produce(int last, Workspace workspace) {
            if (_nextBlockRow >= _rowStarts.Length - 1) return End;
            int row = _nextBlockRow++;
            int end = (int)Math.Min(Height, ((_firstBlockRow + row + 1) << _blockRowShift) - _band.Y0);
            Reserve(end, workspace.Pool);
            int start = Offset(End), length = (end - End) * Width;
            if (Ints != null) Array.Clear(Ints, start, length);
            else Array.Clear(Floats!, start, length);
            CodeBlockDecoder decoder = workspace.Blocks;
            for (int index = _rowStarts[row]; index < _rowStarts[row + 1]; index++) {
                CodeBlock block = _blocks[index];
                if (!decoder.Decode(block, _band, _style, workspace.Source)) continue;
                int left = block.X0 - _band.X0, top = RowIndex(block.Y0 - _band.Y0);
                if (Ints != null) decoder.WriteIntegers(Ints, Width, left, top, _regionShift);
                else decoder.WriteFloats(Floats!, Width, left, top, _regionShift, _band.StepSize);
            }
            return end;
        }
    }

    /// <summary>
    /// One inverse wavelet level (F.3): horizontal synthesis of interleaved rows from the lower level and the HL, LH
    /// and HH subbands, then vertical synthesis over a window that extends past the produced rows by the filter
    /// margin. Every sample is computed with the same operations in the same order as a whole-plane transform.
    /// </summary>
    private sealed class LevelSource : RowSource {
        private readonly RowSource _lower;
        private readonly BandSource _hl, _lh, _hh;
        private readonly int _x0, _y0;
        private readonly int _margin;
        private int[]? _intWindow, _intLine, _intCarry;
        private float[]? _floatWindow, _floatLine, _floatCarry, _floatLow, _floatHigh;
        // Horizontally synthesized rows at the end of the last window, which start the next one.
        private int _carryStart, _carryCount;

        internal LevelSource(RowSource lower, BandSource hl, BandSource lh, BandSource hh, Resolution resolution, bool reversible)
            : base(resolution.Width, resolution.Height, reversible) {
            _lower = lower;
            _hl = hl;
            _lh = lh;
            _hh = hh;
            _x0 = resolution.X0;
            _y0 = resolution.Y0;
            _margin = reversible ? 2 : 4;
            if (lower.Width + hl.Width != Width || lh.Width != lower.Width || hh.Width != hl.Width ||
                hl.Height != lower.Height || hh.Height != lh.Height || lower.Height + lh.Height != Height)
                throw new Jpeg2000Exception("The JPEG 2000 subband geometry is inconsistent.");
        }

        // Low-pass rows sit at even absolute coordinates; these map a level row to its subband row.
        private int LowIndex(int index) => (_y0 & 1) == 0 ? (index + 1) >> 1 : index >> 1;

        private bool IsLow(int index) => ((_y0 + index) & 1) == 0;

        protected override int Produce(int last, Workspace workspace) {
            int first = End;
            int end = Math.Min(Height, Math.Max(last + 1, first + MinimumLevelRows));
            int low = Height == 1 ? 0 : first - _margin, high = Height == 1 ? 1 : end + _margin;
            int carried = _carryCount > 0 && _carryStart == low ? _carryCount : 0;
            int minimumLow = int.MaxValue, maximumLow = -1, minimumHigh = int.MaxValue, maximumHigh = -1;
            for (int position = low + carried; position < high; position++) {
                int index = Height == 1 ? 0 : Reflect(position, Height);
                if (IsLow(index)) {
                    int row = LowIndex(index);
                    minimumLow = Math.Min(minimumLow, row);
                    maximumLow = Math.Max(maximumLow, row);
                } else {
                    int row = index - LowIndex(index);
                    minimumHigh = Math.Min(minimumHigh, row);
                    maximumHigh = Math.Max(maximumHigh, row);
                }
            }
            if (maximumLow >= 0) {
                _lower.Ensure(minimumLow, maximumLow, workspace);
                _hl.Ensure(minimumLow, maximumLow, workspace);
            }
            if (maximumHigh >= 0) {
                _lh.Ensure(minimumHigh, maximumHigh, workspace);
                _hh.Ensure(minimumHigh, maximumHigh, workspace);
            }

            int width = Width, rows = high - low;
            int carry = end < Height ? Math.Min(2 * _margin, rows) : 0;
            long windowLength = (long)rows * width;
            if (windowLength > int.MaxValue) throw new Jpeg2000Exception("The JPEG 2000 tile-component is too wide.");
            SamplePool pool = workspace.Pool;
            if (Reversible) {
                if (_intWindow == null || _intWindow.Length < windowLength) {
                    pool.Return(_intWindow);
                    _intWindow = pool.RentIntegers((int)windowLength);
                }
                if (_intLine == null || _intLine.Length < width + 8) {
                    pool.Return(_intLine);
                    _intLine = pool.RentIntegers(width + 8);
                }
                if (carry > 0 && _intCarry == null) _intCarry = pool.RentIntegers(2 * _margin * width);
            } else {
                if (_floatWindow == null || _floatWindow.Length < windowLength) {
                    pool.Return(_floatWindow);
                    _floatWindow = pool.RentFloats((int)windowLength);
                }
                if (_floatLine == null || _floatLine.Length < width + 8) {
                    pool.Return(_floatLine);
                    _floatLine = pool.RentFloats(width + 8);
                }
                if (carry > 0 && _floatCarry == null) _floatCarry = pool.RentFloats(2 * _margin * width);
#if NET8_0_OR_GREATER
                if (width >= PolyphaseMinimumWidth && Vector.IsHardwareAccelerated && _floatLow == null) {
                    _floatLow = pool.RentFloats(width / 2 + 8);
                    _floatHigh = pool.RentFloats(width / 2 + 8);
                }
#endif
            }

            // HOR_SR on every window row not carried over from the previous window.
            if (carried > 0) {
                if (Reversible) Array.Copy(_intCarry!, 0, _intWindow!, 0, carried * width);
                else Array.Copy(_floatCarry!, 0, _floatWindow!, 0, carried * width);
            }
            for (int position = low + carried; position < high; position++) {
                int index = Height == 1 ? 0 : Reflect(position, Height);
                int target = (position - low) * width;
                bool lowRow = IsLow(index);
                int subbandRow = lowRow ? LowIndex(index) : index - LowIndex(index);
                RowSource left = lowRow ? _lower : _lh, right = lowRow ? _hl : _hh;
                int leftOffset = left.Offset(subbandRow), rightOffset = right.Offset(subbandRow);
                if (Reversible) {
                    int[] line = _intLine!;
                    Interleave(left.Ints!, leftOffset, right.Ints!, rightOffset, width, _x0, line);
                    Synthesize53(line, width, _x0);
                    Array.Copy(line, 4, _intWindow!, target, width);
                } else {
#if NET8_0_OR_GREATER
                    if (_floatLow != null) {
                        SynthesizePolyphase97(left.Floats!, leftOffset, right.Floats!, rightOffset, width, _x0, _floatLow, _floatHigh!, _floatWindow!, target);
                        continue;
                    }
#endif
                    float[] line = _floatLine!;
                    Interleave(left.Floats!, leftOffset, right.Floats!, rightOffset, width, _x0, line);
                    Synthesize97(line, width, _x0);
                    Array.Copy(line, 4, _floatWindow!, target, width);
                }
            }

            _carryCount = carry;
            _carryStart = high - carry;
            if (carry > 0) {
                if (Reversible) Array.Copy(_intWindow!, (rows - carry) * width, _intCarry!, 0, carry * width);
                else Array.Copy(_floatWindow!, (rows - carry) * width, _floatCarry!, 0, carry * width);
            }

            // VER_SR on the window, then keep the rows that have their full filter support.
            if (Reversible) SynthesizeRows53(_intWindow!, width, low, high, Height, _y0);
            else SynthesizeRows97(_floatWindow!, width, low, high, Height, _y0);
            Reserve(end, pool);
            int count = (end - first) * width;
            if (Reversible) Array.Copy(_intWindow!, (first - low) * width, Ints!, Offset(first), count);
            else Array.Copy(_floatWindow!, (first - low) * width, Floats!, Offset(first), count);
            return end;
        }

        internal override void Return(SamplePool pool) {
            base.Return(pool);
            pool.Return(_intWindow);
            pool.Return(_intLine);
            pool.Return(_intCarry);
            pool.Return(_floatWindow);
            pool.Return(_floatLine);
            pool.Return(_floatCarry);
            pool.Return(_floatLow);
            pool.Return(_floatHigh);
            _intWindow = _intLine = _intCarry = null;
            _floatWindow = _floatLine = _floatCarry = _floatLow = _floatHigh = null;
            _carryCount = 0;
            _lower.Return(pool);
            _hl.Return(pool);
            _lh.Return(pool);
            _hh.Return(pool);
        }
    }

    // 2D_INTERLEAVE along one row: low-pass samples sit at even absolute coordinates.
    private static void Interleave<T>(T[] lowSource, int lowStart, T[] highSource, int highStart, int count, int origin, T[] line) {
        int firstLow = origin & 1;
        for (int index = 4 + firstLow, source = lowStart; index < 4 + count; index += 2, source++) line[index] = lowSource[source];
        for (int index = 5 - firstLow, source = highStart; index < 4 + count; index += 2, source++) line[index] = highSource[source];
    }

    // Periodic symmetric extension (F.3.7).
    private static int Reflect(int index, int count) {
        int period = 2 * (count - 1);
        index %= period;
        if (index < 0) index += period;
        return index >= count ? period - index : index;
    }

    private static void Extend<T>(T[] line, int count, int padding) {
        for (int k = 1; k <= padding; k++) {
            line[4 - k] = line[4 + Reflect(-k, count)];
            line[4 + count - 1 + k] = line[4 + Reflect(count - 1 + k, count)];
        }
    }

    // 1D_SR for the reversible 5/3 filter (F.3.8.1).
    private static void Synthesize53(int[] line, int count, int origin) {
        if (count == 1) {
            if ((origin & 1) != 0) line[4] /= 2;
            return;
        }
        Extend(line, count, 2);
        int first = ((origin - 1) & 1) == 0 ? -1 : 0;
        for (int j = first; j <= count; j += 2) line[4 + j] -= (line[3 + j] + line[5 + j] + 2) >> 2;
        first = (origin & 1) != 0 ? 0 : 1;
        for (int j = first; j < count; j += 2) line[4 + j] += (line[3 + j] + line[5 + j]) >> 1;
    }

    // 1D_SR for the irreversible 9/7 filter (F.3.8.2).
    private static void Synthesize97(float[] line, int count, int origin) {
        if (count == 1) {
            if ((origin & 1) != 0) line[4] *= 0.5F;
            return;
        }
        Extend(line, count, 4);
        int evenStart = ((origin - 4) & 1) == 0 ? -4 : -3;
        int oddStart = evenStart == -4 ? -3 : -4;
        const float inverseK = 1F / IrreversibleK;
        for (int j = evenStart; j < count + 4; j += 2) line[4 + j] *= IrreversibleK;
        for (int j = oddStart; j < count + 4; j += 2) line[4 + j] *= inverseK;
        int even = ((origin - 3) & 1) == 0 ? -3 : -2;
        for (int j = even; j < count + 3; j += 2) line[4 + j] -= IrreversibleDelta * (line[3 + j] + line[5 + j]);
        int odd = ((origin - 2) & 1) != 0 ? -2 : -1;
        for (int j = odd; j < count + 2; j += 2) line[4 + j] -= IrreversibleGamma * (line[3 + j] + line[5 + j]);
        even = ((origin - 1) & 1) == 0 ? -1 : 0;
        for (int j = even; j < count + 1; j += 2) line[4 + j] -= IrreversibleBeta * (line[3 + j] + line[5 + j]);
        odd = (origin & 1) != 0 ? 0 : 1;
        for (int j = odd; j < count; j += 2) line[4 + j] -= IrreversibleAlpha * (line[3 + j] + line[5 + j]);
    }

#if NET8_0_OR_GREATER
    /// <summary>Rows at least this wide run the horizontal 9/7 synthesis in polyphase form.</summary>
    private const int PolyphaseMinimumWidth = 32;

    // Synthesize97 on one row with the low-pass (even absolute) and high-pass positions of the extended line kept
    // in separate arrays, so each lifting step reads its neighbours contiguously and runs on vectors. Every
    // position is updated by the same steps, with the same operands, as in Synthesize97.
    private static void SynthesizePolyphase97(float[] lowSource, int lowStart, float[] highSource, int highStart, int count, int origin,
        float[] low, float[] high, float[] target, int targetStart) {
        int lowBase = (origin & 1) == 0 ? -4 : -3; // first low-pass position of the extended line [-4, count + 4)
        int highBase = lowBase == -4 ? -3 : -4;
        int firstLow = lowBase + 4, firstHigh = highBase + 4; // first positions inside [0, count)
        int lowCount = (count - firstLow + 1) >> 1, highCount = count - lowCount;
        Array.Copy(lowSource, lowStart, low, (firstLow - lowBase) >> 1, lowCount);
        Array.Copy(highSource, highStart, high, (firstHigh - highBase) >> 1, highCount);
        for (int k = 1; k <= 4; k++) {
            ExtendPolyphase(-k);
            ExtendPolyphase(count - 1 + k);
        }
        int lowLength = (count + 4 - lowBase + 1) >> 1, highLength = (count + 4 - highBase + 1) >> 1;
        const float inverseK = 1F / IrreversibleK;
        ScaleSpan(low, 0, lowLength, IrreversibleK);
        ScaleSpan(high, 0, highLength, inverseK);
        // Neighbours of low position lowBase + 2k are high entries k + lowShift and k + lowShift + 1, and the
        // reverse for high positions.
        int lowShift = highBase == lowBase + 1 ? -1 : 0, highShift = lowBase == highBase + 1 ? -1 : 0;
        LiftSpan(low, lowBase, ((origin - 3) & 1) == 0 ? -3 : -2, count + 3, high, lowShift, IrreversibleDelta);
        LiftSpan(high, highBase, ((origin - 2) & 1) != 0 ? -2 : -1, count + 2, low, highShift, IrreversibleGamma);
        LiftSpan(low, lowBase, ((origin - 1) & 1) == 0 ? -1 : 0, count + 1, high, lowShift, IrreversibleBeta);
        LiftSpan(high, highBase, (origin & 1) != 0 ? 0 : 1, count, low, highShift, IrreversibleAlpha);
        // Positions 0 and 1 hold one sample of each phase; whole vectors of pairs are interleaved first.
        int lowIndex = (firstLow - lowBase) >> 1, highIndex = (firstHigh - highBase) >> 1;
        int paired = firstLow == 0
            ? InterleavePairs(low, lowIndex, high, highIndex, target, targetStart, count >> 1)
            : InterleavePairs(high, highIndex, low, lowIndex, target, targetStart, count >> 1);
        for (int position = firstLow + 2 * paired, k = lowIndex + paired; position < count; position += 2, k++) target[targetStart + position] = low[k];
        for (int position = firstHigh + 2 * paired, k = highIndex + paired; position < count; position += 2, k++) target[targetStart + position] = high[k];

        // Periodic symmetric extension (F.3.7) keeps the parity of a position, so each side extends from itself.
        void ExtendPolyphase(int position) {
            int source = Reflect(position, count);
            bool lowPosition = ((origin + position) & 1) == 0;
            float value = lowPosition ? lowSource[lowStart + ((source - firstLow) >> 1)] : highSource[highStart + ((source - firstHigh) >> 1)];
            if (lowPosition) low[(position - lowBase) >> 1] = value;
            else high[(position - highBase) >> 1] = value;
        }
    }

    // target[start + 2i] = even[evenStart + i], target[start + 2i + 1] = odd[oddStart + i] for whole vectors of i below
    // `pairs`; returns how many pairs were written. A pure copy, so the samples are bit-identical.
    private static int InterleavePairs(float[] even, int evenStart, float[] odd, int oddStart, float[] target, int targetStart, int pairs) {
        int lanes = Vector128<float>.Count, done = 0;
        if (pairs < lanes || evenStart + pairs > even.Length || oddStart + pairs > odd.Length || targetStart + 2 * pairs > target.Length) return 0;
        if (AdvSimd.Arm64.IsSupported) {
            for (; done <= pairs - lanes; done += lanes) {
                Vector128<float> a = Vector128.Create(even.AsSpan(evenStart + done, lanes));
                Vector128<float> b = Vector128.Create(odd.AsSpan(oddStart + done, lanes));
                AdvSimd.Arm64.ZipLow(a, b).CopyTo(target, targetStart + 2 * done);
                AdvSimd.Arm64.ZipHigh(a, b).CopyTo(target, targetStart + 2 * done + lanes);
            }
        } else if (Sse.IsSupported) {
            for (; done <= pairs - lanes; done += lanes) {
                Vector128<float> a = Vector128.Create(even.AsSpan(evenStart + done, lanes));
                Vector128<float> b = Vector128.Create(odd.AsSpan(oddStart + done, lanes));
                Sse.UnpackLow(a, b).CopyTo(target, targetStart + 2 * done);
                Sse.UnpackHigh(a, b).CopyTo(target, targetStart + 2 * done + lanes);
            }
        }
        return done;
    }

    private static void ScaleSpan(float[] values, int start, int end, float factor) {
        int index = start;
        var scale = new Vector<float>(factor);
        for (int lanes = Vector<float>.Count; index <= end - lanes; index += lanes) (new Vector<float>(values, index) * scale).CopyTo(values, index);
        for (; index < end; index++) values[index] *= factor;
    }

    // values[k] -= factor * (neighbours[k + shift] + neighbours[k + shift + 1]) for positions [from, to) of one parity.
    private static void LiftSpan(float[] values, int valuesBase, int from, int to, float[] neighbours, int shift, float factor) {
        int index = (from - valuesBase) >> 1, end = index + ((to - from + 1) >> 1);
        var weight = new Vector<float>(factor);
        for (int lanes = Vector<float>.Count; index <= end - lanes; index += lanes) {
            var sum = new Vector<float>(neighbours, index + shift) + new Vector<float>(neighbours, index + shift + 1);
            (new Vector<float>(values, index) - weight * sum).CopyTo(values, index);
        }
        for (; index < end; index++) values[index] -= factor * (neighbours[index + shift] + neighbours[index + shift + 1]);
    }
#endif

    // First position at or after `from` with the parity of `start`.
    private static int Align(int from, int start) => from + ((from - start) & 1);

    // VER_SR for 5/3 over window rows holding positions [low, high) of a column of `count` samples. Each lifting
    // step runs over the positions a whole-column transform would update, limited to those with both neighbours in
    // the window; positions [low + 2, high - 2) therefore come out exactly as from the whole column.
    private static void SynthesizeRows53(int[] rows, int width, int low, int high, int count, int origin) {
        if (count == 1) {
            if ((origin & 1) != 0) {
                for (int x = 0; x < width; x++) rows[x] /= 2;
            }
            return;
        }
        int first = ((origin - 1) & 1) == 0 ? -1 : 0;
        int stop = Math.Min(count + 1, high - 1);
        for (int j = Align(Math.Max(first, low + 1), first); j < stop; j += 2) {
            int center = (j - low) * width, above = center - width, below = center + width;
            int x = 0;
#if NET8_0_OR_GREATER
            if (Vector.IsHardwareAccelerated) {
                var two = new Vector<int>(2);
                for (int lanes = Vector<int>.Count; x <= width - lanes; x += lanes) {
                    var sum = new Vector<int>(rows, above + x) + new Vector<int>(rows, below + x) + two;
                    (new Vector<int>(rows, center + x) - Vector.ShiftRightArithmetic(sum, 2)).CopyTo(rows, center + x);
                }
            }
#endif
            for (; x < width; x++) rows[center + x] -= (rows[above + x] + rows[below + x] + 2) >> 2;
        }
        first = (origin & 1) != 0 ? 0 : 1;
        stop = Math.Min(count, high - 1);
        for (int j = Align(Math.Max(first, low + 1), first); j < stop; j += 2) {
            int center = (j - low) * width, above = center - width, below = center + width;
            int x = 0;
#if NET8_0_OR_GREATER
            if (Vector.IsHardwareAccelerated) {
                for (int lanes = Vector<int>.Count; x <= width - lanes; x += lanes) {
                    var sum = new Vector<int>(rows, above + x) + new Vector<int>(rows, below + x);
                    (new Vector<int>(rows, center + x) + Vector.ShiftRightArithmetic(sum, 1)).CopyTo(rows, center + x);
                }
            }
#endif
            for (; x < width; x++) rows[center + x] += (rows[above + x] + rows[below + x]) >> 1;
        }
    }

    // VER_SR for 9/7 over window rows holding positions [low, high); see SynthesizeRows53.
    private static void SynthesizeRows97(float[] rows, int width, int low, int high, int count, int origin) {
        if (count == 1) {
            if ((origin & 1) != 0) {
                for (int x = 0; x < width; x++) rows[x] *= 0.5F;
            }
            return;
        }
        int evenStart = ((origin - 4) & 1) == 0 ? -4 : -3;
        int oddStart = evenStart == -4 ? -3 : -4;
        const float inverseK = 1F / IrreversibleK;
        ScaleRows(rows, width, low, high, evenStart, count + 4, IrreversibleK);
        ScaleRows(rows, width, low, high, oddStart, count + 4, inverseK);
        LiftRows(rows, width, low, high, ((origin - 3) & 1) == 0 ? -3 : -2, count + 3, IrreversibleDelta);
        LiftRows(rows, width, low, high, ((origin - 2) & 1) != 0 ? -2 : -1, count + 2, IrreversibleGamma);
        LiftRows(rows, width, low, high, ((origin - 1) & 1) == 0 ? -1 : 0, count + 1, IrreversibleBeta);
        LiftRows(rows, width, low, high, (origin & 1) != 0 ? 0 : 1, count, IrreversibleAlpha);
    }

    private static void ScaleRows(float[] rows, int width, int low, int high, int start, int stop, float factor) {
        stop = Math.Min(stop, high);
        for (int j = Align(Math.Max(start, low), start); j < stop; j += 2) {
            int center = (j - low) * width;
            int x = 0;
#if NET8_0_OR_GREATER
            if (Vector.IsHardwareAccelerated) {
                var scale = new Vector<float>(factor);
                for (int lanes = Vector<float>.Count; x <= width - lanes; x += lanes) (new Vector<float>(rows, center + x) * scale).CopyTo(rows, center + x);
            }
#endif
            for (; x < width; x++) rows[center + x] *= factor;
        }
    }

    private static void LiftRows(float[] rows, int width, int low, int high, int start, int stop, float factor) {
        stop = Math.Min(stop, high - 1);
        for (int j = Align(Math.Max(start, low + 1), start); j < stop; j += 2) {
            int center = (j - low) * width, above = center - width, below = center + width;
            int x = 0;
#if NET8_0_OR_GREATER
            // Lane-wise single-precision add, multiply and subtract: the same rounding as the scalar loop.
            if (Vector.IsHardwareAccelerated) {
                var weight = new Vector<float>(factor);
                for (int lanes = Vector<float>.Count; x <= width - lanes; x += lanes) {
                    var sum = new Vector<float>(rows, above + x) + new Vector<float>(rows, below + x);
                    (new Vector<float>(rows, center + x) - weight * sum).CopyTo(rows, center + x);
                }
            }
#endif
            for (; x < width; x++) rows[center + x] -= factor * (rows[above + x] + rows[below + x]);
        }
    }
}
