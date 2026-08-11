namespace OfficeIMO.Pdf;

internal static partial class PdfColorSpaceFunctionResolver {
    private static bool TryCreateSampledEvaluator(
        int order,
        double[] domain,
        double[] encode,
        int[] sizes,
        int bitsPerSample,
        ulong maximumSample,
        double[] decode,
        byte[] samples,
        int outputCount,
        out Func<double[], double[]?> evaluator,
        out int cubicEvaluationCost) {
        evaluator = null!;
        cubicEvaluationCost = 0;
        SampleInterpolationBounds interpolationBounds = CreateSampleInterpolationBounds(order, encode, sizes);
        bool useCubic = interpolationBounds.HasCubicDimension;
        long[] strides = CreateSampleStrides(sizes);
        double[]? secondDerivatives = null;
        if (useCubic && sizes.Length == 1 &&
            !TryCreateNaturalSplineSecondDerivatives(
                interpolationBounds.GetStart(0),
                interpolationBounds.GetEnd(0),
                outputCount,
                bitsPerSample,
                maximumSample,
                decode,
                samples,
                out secondDerivatives)) return false;

        evaluator = values => useCubic
            ? EvaluateCubicSampled(values, domain, encode, sizes, interpolationBounds, strides, bitsPerSample, maximumSample, decode, samples, outputCount, secondDerivatives)
            : EvaluateLinearSampled(values, domain, encode, sizes, strides, bitsPerSample, maximumSample, decode, samples, outputCount);
        if (useCubic) cubicEvaluationCost = CalculateCubicEvaluationCost(interpolationBounds, outputCount);
        return true;
    }

    private static double[]? EvaluateLinearSampled(
        double[] values,
        double[] domain,
        double[] encode,
        int[] sizes,
        long[] strides,
        int bitsPerSample,
        ulong maximumSample,
        double[] decode,
        byte[] samples,
        int outputCount) {
        if (!TryCreateSampleCoordinates(values, domain, encode, sizes, interpolationBounds: null, out SampleCoordinates coordinates)) return null;

        var result = new double[outputCount];
        int cornerCount = 1 << sizes.Length;
        for (int corner = 0; corner < cornerCount; corner++) {
            double weight = 1D;
            long pointIndex = 0;
            for (int input = 0; input < sizes.Length; input++) {
                int lower = coordinates.GetLower(input);
                int upper = Math.Min(lower + 1, sizes[input] - 1);
                bool useUpper = (corner & (1 << input)) != 0;
                int coordinate = useUpper ? upper : lower;
                double fraction = coordinates.GetFraction(input);
                weight *= useUpper ? fraction : 1D - fraction;
                pointIndex += coordinate * strides[input];
            }
            if (weight == 0D) continue;

            long sampleOffset = pointIndex * outputCount;
            for (int output = 0; output < outputCount; output++) {
                result[output] += weight * ReadDecodedSample(
                    samples,
                    sampleOffset + output,
                    bitsPerSample,
                    maximumSample,
                    decode,
                    output);
            }
        }
        return result;
    }

    private static double[]? EvaluateCubicSampled(
        double[] values,
        double[] domain,
        double[] encode,
        int[] sizes,
        SampleInterpolationBounds interpolationBounds,
        long[] strides,
        int bitsPerSample,
        ulong maximumSample,
        double[] decode,
        byte[] samples,
        int outputCount,
        double[]? secondDerivatives) {
        if (!TryCreateSampleCoordinates(values, domain, encode, sizes, interpolationBounds, out SampleCoordinates coordinates)) return null;

        var result = new double[outputCount];
        if (sizes.Length == 1) {
            if (secondDerivatives == null) return null;
            int left = coordinates.GetLower(0);
            double fraction = coordinates.GetFraction(0);
            int splineStart = interpolationBounds.GetStart(0);
            int splinePointCount = interpolationBounds.GetEnd(0) - splineStart + 1;
            for (int output = 0; output < outputCount; output++) {
                double leftValue = ReadDecodedSample(samples, (long)left * outputCount + output, bitsPerSample, maximumSample, decode, output);
                double rightValue = ReadDecodedSample(samples, (long)(left + 1) * outputCount + output, bitsPerSample, maximumSample, decode, output);
                int derivativeOffset = output * splinePointCount + left - splineStart;
                double leftSecond = secondDerivatives[derivativeOffset];
                double rightSecond = secondDerivatives[derivativeOffset + 1];
                double complement = 1D - fraction;
                result[output] = leftValue * complement + rightValue * fraction -
                    fraction * complement *
                    (leftSecond * (complement + 1D) + rightSecond * (fraction + 1D)) / 6D;
                if (!IsFinite(result[output])) return null;
            }
            return result;
        }

        for (int output = 0; output < outputCount; output++) {
            result[output] = InterpolateCubicTensor(
                sizes.Length - 1,
                0L,
                coordinates,
                interpolationBounds,
                strides,
                sizes,
                samples,
                bitsPerSample,
                maximumSample,
                decode,
                outputCount,
                output);
            if (!IsFinite(result[output])) return null;
        }
        return result;
    }

    private static double InterpolateCubicTensor(
        int dimension,
        long pointIndex,
        SampleCoordinates coordinates,
        SampleInterpolationBounds interpolationBounds,
        long[] strides,
        int[] sizes,
        byte[] samples,
        int bitsPerSample,
        ulong maximumSample,
        double[] decode,
        int outputCount,
        int output) {
        if (dimension < 0) {
            return ReadDecodedSample(
                samples,
                pointIndex * outputCount + output,
                bitsPerSample,
                maximumSample,
                decode,
                output);
        }

        long step = strides[dimension];
        int coordinate = coordinates.GetLower(dimension);
        double value1 = InterpolateCubicTensor(dimension - 1, pointIndex + coordinate * step, coordinates, interpolationBounds, strides, sizes, samples, bitsPerSample, maximumSample, decode, outputCount, output);
        if (interpolationBounds.GetStart(dimension) == interpolationBounds.GetEnd(dimension)) return value1;
        double value2 = InterpolateCubicTensor(dimension - 1, pointIndex + (coordinate + 1L) * step, coordinates, interpolationBounds, strides, sizes, samples, bitsPerSample, maximumSample, decode, outputCount, output);
        double fraction = coordinates.GetFraction(dimension);
        if (!interpolationBounds.UsesCubic(dimension)) return value1 * (1D - fraction) + value2 * fraction;
        double value0 = coordinate > interpolationBounds.GetStart(dimension)
            ? InterpolateCubicTensor(dimension - 1, pointIndex + (coordinate - 1L) * step, coordinates, interpolationBounds, strides, sizes, samples, bitsPerSample, maximumSample, decode, outputCount, output)
            : 2D * value1 - value2;
        double value3 = coordinate + 2 <= interpolationBounds.GetEnd(dimension)
            ? InterpolateCubicTensor(dimension - 1, pointIndex + (coordinate + 2L) * step, coordinates, interpolationBounds, strides, sizes, samples, bitsPerSample, maximumSample, decode, outputCount, output)
            : 2D * value2 - value1;
        return value1 + 0.5D * fraction *
            (value2 - value0 + fraction *
                (2D * value0 - 5D * value1 + 4D * value2 - value3 +
                    fraction * (3D * (value1 - value2) + value3 - value0)));
    }

    private static bool TryCreateNaturalSplineSecondDerivatives(
        int start,
        int end,
        int outputCount,
        int bitsPerSample,
        ulong maximumSample,
        double[] decode,
        byte[] samples,
        out double[] derivatives) {
        int pointCount = checked(end - start + 1);
        derivatives = new double[checked(pointCount * outputCount)];
        var workspace = new double[pointCount];
        for (int output = 0; output < outputCount; output++) {
            int offset = output * pointCount;
            derivatives[offset] = 0D;
            workspace[0] = 0D;
            for (int sample = 1; sample < pointCount - 1; sample++) {
                double denominator = 0.5D * derivatives[offset + sample - 1] + 2D;
                derivatives[offset + sample] = -0.5D / denominator;
                double previous = ReadDecodedSample(samples, (long)(start + sample - 1) * outputCount + output, bitsPerSample, maximumSample, decode, output);
                double current = ReadDecodedSample(samples, (long)(start + sample) * outputCount + output, bitsPerSample, maximumSample, decode, output);
                double next = ReadDecodedSample(samples, (long)(start + sample + 1) * outputCount + output, bitsPerSample, maximumSample, decode, output);
                double curvature = 3D * (previous - 2D * current + next);
                workspace[sample] = (curvature - 0.5D * workspace[sample - 1]) / denominator;
                if (!IsFinite(workspace[sample])) return false;
            }
            derivatives[offset + pointCount - 1] = 0D;
            for (int sample = pointCount - 2; sample >= 0; sample--) {
                derivatives[offset + sample] = derivatives[offset + sample] * derivatives[offset + sample + 1] + workspace[sample];
                if (!IsFinite(derivatives[offset + sample])) return false;
            }
        }
        return true;
    }

    private static long[] CreateSampleStrides(int[] sizes) {
        var strides = new long[sizes.Length];
        long stride = 1L;
        for (int input = 0; input < sizes.Length; input++) {
            strides[input] = stride;
            stride = checked(stride * sizes[input]);
        }
        return strides;
    }

    private static int CalculateCubicEvaluationCost(SampleInterpolationBounds interpolationBounds, int outputCount) {
        int leafReads = 1;
        for (int input = 0; input < interpolationBounds.InputCount; input++) {
            int branchCount = interpolationBounds.UsesCubic(input)
                ? 4
                : interpolationBounds.GetStart(input) == interpolationBounds.GetEnd(input) ? 1 : 2;
            leafReads = checked(leafReads * branchCount);
        }
        return checked(leafReads * outputCount);
    }

    private static int GetNaturalSplinePointCount(int order, int inputCount, double[] encode, int[] sizes) {
        return order == 3 && inputCount == 1 && sizes[0] >= 4 ? sizes[0] : 0;
    }

    private static SampleInterpolationBounds CreateSampleInterpolationBounds(int order, double[] encode, int[] sizes) {
        var starts = new int[sizes.Length];
        var ends = new int[sizes.Length];
        var cubic = new bool[sizes.Length];
        bool hasCubic = false;
        for (int input = 0; input < sizes.Length; input++) {
            starts[input] = 0;
            ends[input] = sizes[input] - 1;
            cubic[input] = order == 3 && sizes[input] >= 4;
            hasCubic |= cubic[input];
        }
        return new SampleInterpolationBounds(starts, ends, cubic, hasCubic);
    }

    private static bool TryEncodeSampleCoordinate(
        double value,
        double[] domain,
        double[] encode,
        int[] sizes,
        int input,
        out double encoded) {
        encoded = PdfColorFunction.Interpolate(
            value,
            domain[input * 2],
            domain[input * 2 + 1],
            encode[input * 2],
            encode[input * 2 + 1]);
        if (!IsFinite(encoded)) return false;
        encoded = PdfColorFunction.Clamp(encoded, 0D, sizes[input] - 1D);
        return true;
    }

    private static bool TryCreateSampleCoordinates(
        double[] values,
        double[] domain,
        double[] encode,
        int[] sizes,
        SampleInterpolationBounds? interpolationBounds,
        out SampleCoordinates coordinates) {
        coordinates = default;
        for (int input = 0; input < sizes.Length; input++) {
            if (!TryEncodeSampleCoordinate(values[input], domain, encode, sizes, input, out double encoded)) return false;
            int lower = (int)Math.Floor(encoded);
            if (interpolationBounds != null && interpolationBounds.GetStart(input) < interpolationBounds.GetEnd(input)) {
                lower = Math.Min(lower, interpolationBounds.GetEnd(input) - 1);
            }
            coordinates.Set(input, lower, encoded - lower);
        }
        return true;
    }

    private static double ReadDecodedSample(
        byte[] samples,
        long sampleIndex,
        int bitsPerSample,
        ulong maximumSample,
        double[] decode,
        int output) {
        ulong raw = ReadSample(samples, sampleIndex, bitsPerSample);
        return PdfColorFunction.Interpolate(raw, 0D, maximumSample, decode[output * 2], decode[output * 2 + 1]);
    }

    private static ulong ReadSample(byte[] samples, long sampleIndex, int bitsPerSample) {
        long bitOffset = sampleIndex * bitsPerSample;
        if ((bitOffset & 7L) == 0L && bitsPerSample is 8 or 16 or 24 or 32) {
            int byteIndex = checked((int)(bitOffset >> 3));
            ulong alignedValue = 0UL;
            int byteCount = bitsPerSample / 8;
            for (int index = 0; index < byteCount; index++) alignedValue = (alignedValue << 8) | samples[byteIndex + index];
            return alignedValue;
        }

        ulong value = 0UL;
        for (int bit = 0; bit < bitsPerSample; bit++) {
            long absoluteBit = bitOffset + bit;
            int byteIndex = checked((int)(absoluteBit >> 3));
            int bitInByte = 7 - (int)(absoluteBit & 7L);
            value = (value << 1) | (uint)((samples[byteIndex] >> bitInByte) & 1);
        }
        return value;
    }

    private static double[] CreateSampleBreakpoints(double[] domain, double[] encode, int[] sizes, int order) {
        if (sizes.Length != 1 || encode[0] == encode[1]) return domain.Distinct().ToArray();
        double encodedMinimum = Math.Max(0D, Math.Min(encode[0], encode[1]));
        double encodedMaximum = Math.Min(sizes[0] - 1D, Math.Max(encode[0], encode[1]));
        if (encodedMinimum > encodedMaximum) return domain.Distinct().ToArray();

        var result = new List<double>(MaxSuggestedSampleBreakpoints + 2);
        result.AddRange(domain);
        if (GetNaturalSplinePointCount(order, 1, encode, sizes) > 0) {
            double encodedSpan = encodedMaximum - encodedMinimum;
            int intervalCount = Math.Max(1, (int)Math.Ceiling(encodedSpan));
            int pointCount = Math.Min(MaxSuggestedSampleBreakpoints, checked(intervalCount * 4 + 1));
            for (int point = 0; point < pointCount; point++) {
                double sample = pointCount == 1
                    ? encodedMinimum
                    : encodedMinimum + point * encodedSpan / (pointCount - 1D);
                AddMappedSampleBreakpoint(result, sample, domain, encode);
            }
        } else {
            int firstSample = (int)Math.Ceiling(encodedMinimum);
            int lastSample = (int)Math.Floor(encodedMaximum);
            if (firstSample <= lastSample) {
                int reachablePointCount = checked(lastSample - firstSample + 1);
                int pointCount = Math.Min(reachablePointCount, MaxSuggestedSampleBreakpoints);
                for (int point = 0; point < pointCount; point++) {
                    double sample = pointCount == 1
                        ? firstSample
                        : firstSample + point * (reachablePointCount - 1D) / (pointCount - 1D);
                    AddMappedSampleBreakpoint(result, sample, domain, encode);
                }
            }
        }
        return LimitSuggestedPoints(result, domain);
    }

    private static void AddMappedSampleBreakpoint(List<double> result, double sample, double[] domain, double[] encode) {
        double input = PdfColorFunction.Interpolate(sample, encode[0], encode[1], domain[0], domain[1]);
        if (IsFinite(input) && input >= domain[0] && input <= domain[1]) result.Add(input);
    }

    private struct SampleCoordinates {
        private int _lower0;
        private int _lower1;
        private int _lower2;
        private int _lower3;
        private double _fraction0;
        private double _fraction1;
        private double _fraction2;
        private double _fraction3;

        internal int GetLower(int input) => input switch {
            0 => _lower0,
            1 => _lower1,
            2 => _lower2,
            _ => _lower3
        };

        internal double GetFraction(int input) => input switch {
            0 => _fraction0,
            1 => _fraction1,
            2 => _fraction2,
            _ => _fraction3
        };

        internal void Set(int input, int lower, double fraction) {
            switch (input) {
                case 0:
                    _lower0 = lower;
                    _fraction0 = fraction;
                    break;
                case 1:
                    _lower1 = lower;
                    _fraction1 = fraction;
                    break;
                case 2:
                    _lower2 = lower;
                    _fraction2 = fraction;
                    break;
                default:
                    _lower3 = lower;
                    _fraction3 = fraction;
                    break;
            }
        }
    }

    private sealed class SampleInterpolationBounds {
        private readonly int[] _starts;
        private readonly int[] _ends;
        private readonly bool[] _cubic;

        internal SampleInterpolationBounds(int[] starts, int[] ends, bool[] cubic, bool hasCubicDimension) {
            _starts = starts;
            _ends = ends;
            _cubic = cubic;
            HasCubicDimension = hasCubicDimension;
        }

        internal bool HasCubicDimension { get; }

        internal int InputCount => _starts.Length;

        internal int GetStart(int input) => _starts[input];

        internal int GetEnd(int input) => _ends[input];

        internal bool UsesCubic(int input) => _cubic[input];
    }
}
