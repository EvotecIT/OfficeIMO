namespace OfficeIMO.DjVu;

// Appendix 1's coefficient bands and four decoding passes. Contexts and precision
// survive successive chunks; each chunk has a fresh arithmetic interval.
internal sealed class Iw44Plane {
    private static readonly int[] FirstBucket = { 0, 1, 2, 3, 4, 8, 12, 16, 32, 48 };
    private static readonly int[] BucketCount = { 1, 1, 1, 1, 4, 4, 4, 16, 16, 16 };
    private static readonly int[] StepByCoefficient = CreateStepIndices();
    private readonly int[] _steps = { 0x4000, 0x8000, 0x8000, 0x10000, 0x10000, 0x10000, 0x20000, 0x20000,
        0x20000, 0x40000, 0x40000, 0x40000, 0x80000, 0x40000, 0x40000, 0x80000 };
    private readonly byte[] _bucketContexts = new byte[80], _activationContexts = new byte[16];
    private byte _bandContext, _mantissaContext;
    private readonly byte[] _flags = new byte[1024], _bucketFlags = new byte[64], _decodeBucket = new byte[64];
    internal readonly int[] Coefficients;
    private readonly DjVuReadBudget _budget;
    private int _band;
    internal Iw44Plane(int blocks, DjVuReadBudget budget) { Coefficients = new int[checked(blocks * 1024)]; _budget = budget; }

    internal long CoefficientSamples(int slices) {
        long perBlock = 0;
        for (int i = 0; i < slices; i++) perBlock += BucketCount[(_band + i) % 10] * 16L;
        return perBlock * (Coefficients.LongLength / 1024);
    }

    internal int ReconstructionCoefficient(int index) {
        int value = Coefficients[index];
        int step = _steps[StepByCoefficient[index % 1024]];
        // Newly significant coefficients use a lower reconstruction estimate until
        // their first refinement. Keep the arithmetic decoder's midpoint unchanged.
        // The quarter-step estimate is independently checked by inverse-lifting
        // reference pixels, including progressive and high-frequency fixtures.
        if (value != 0 && Math.Abs(value) <= step * 3) value -= Math.Sign(value) * (step >> 2);
        return value;
    }

    internal void Slice(ZpDecoder arithmetic) {
        int first = FirstBucket[_band], buckets = BucketCount[_band];
        for (int block = 0; block < Coefficients.Length; block += 1024) {
            _budget.Cancellation.ThrowIfCancellationRequested();
            int bandFlags = 0;
            for (int bucket = first; bucket < first + buckets; bucket++) {
                int flags = 0;
                _decodeBucket[bucket] = 0;
                for (int i = bucket * 16; i < bucket * 16 + 16; i++) {
                    int step = _steps[StepByCoefficient[i]];
                    byte flag = step > 0 && step < 0x8000 ? (byte)(Coefficients[block + i] == 0 ? 2 : 1) : (byte)0;
                    _flags[i] = flag; flags |= flag;
                }
                _bucketFlags[bucket] = (byte)flags; bandFlags |= flags;
            }
            bool decode = buckets < 16 || (bandFlags & 1) != 0 || (bandFlags & 2) != 0 && arithmetic.Bit(ref _bandContext) != 0;
            if (decode) {
                for (int bucket = first; bucket < first + buckets; bucket++) {
                    if ((_bucketFlags[bucket] & 2) == 0) continue;
                    int parents = 0;
                    if (_band != 0) for (int i = bucket * 4; i < bucket * 4 + 4; i++) if (Coefficients[block + i] != 0) parents++;
                    int context = _band * 8 + Math.Min(parents, 3) + ((bandFlags & 1) != 0 ? 4 : 0);
                    _decodeBucket[bucket] = (byte)arithmetic.Bit(ref _bucketContexts[context]);
                }
                for (int bucket = first; bucket < first + buckets; bucket++) {
                    if (_decodeBucket[bucket] == 0) continue;
                    int potential = 0;
                    for (int i = bucket * 16; i < bucket * 16 + 16; i++) if (_flags[i] == 2) potential++;
                    for (int i = bucket * 16; i < bucket * 16 + 16; i++) {
                        if (_flags[i] != 2) continue;
                        int context = Math.Min(potential, 7) + ((_bucketFlags[bucket] & 1) != 0 ? 8 : 0);
                        if (arithmetic.Bit(ref _activationContexts[context]) != 0) {
                            int step = _steps[StepByCoefficient[i]];
                            int value = step + (step >> 1);
                            Coefficients[block + i] = arithmetic.WaveletBit() != 0 ? -value : value;
                            potential = 0;
                        }
                        if (potential > 0) potential--;
                    }
                }
            }
            for (int i = first * 16; i < (first + buckets) * 16; i++) {
                if (_flags[i] != 1) continue;
                int step = _steps[StepByCoefficient[i]];
                int value = Coefficients[block + i], magnitude = Math.Abs(value);
                int increase = magnitude <= step * 3 ? arithmetic.Bit(ref _mantissaContext) : arithmetic.WaveletBit();
                magnitude += increase != 0 ? step >> 1 : -(step == 1 ? 1 : step >> 1);
                Coefficients[block + i] = value < 0 ? -magnitude : magnitude;
            }
        }
        if (_band == 0) for (int i = 0; i < 7; i++) _steps[i] >>= 1;
        else _steps[_band + 6] >>= 1;
        _band = (_band + 1) % 10;
    }

    private static int[] CreateStepIndices() {
        var result = new int[1024];
        for (int i = 0; i < result.Length; i++) {
            result[i] = i < 4 ? i : i < 16 ? 4 + (i - 4) / 4 : i < 64 ? 7 + (i - 16) / 16
                : i < 256 ? 10 + (i - 64) / 64 : 13 + (i - 256) / 256;
        }
        return result;
    }
}
