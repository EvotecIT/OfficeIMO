namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // Row-oriented form of the eight-bit SampleYccToRgb path for full-resolution luma
    // and chroma subsampled by one or two in each direction. Interpolated chroma is an
    // exact multiple of 1/16 there, so it is carried as an integer in sixteenths and the
    // floating-point terms come from tables holding the same IEEE products the per-pixel
    // path computes. Output is identical to SampleYccToRgb for every pixel.
    private static bool TryComposeYccHighQuality8(
        byte[] output,
        int width,
        int height,
        BaselineComponentState luma,
        BaselineComponentState cb,
        BaselineComponentState cr,
        int maximumHorizontalSampling,
        int maximumVerticalSampling,
        bool outputRgba,
        CancellationToken cancellationToken) {
        if (luma.WideBuffer != null || cb.WideBuffer != null || cr.WideBuffer != null) return false;
        if (luma.Component.H != maximumHorizontalSampling || luma.Component.V != maximumVerticalSampling) return false;
        if (!TryCreateChromaAxis(cb.Component.H, maximumHorizontalSampling, width, out ChromaAxis cbColumns) ||
            !TryCreateChromaAxis(cr.Component.H, maximumHorizontalSampling, width, out ChromaAxis crColumns) ||
            !TryCreateChromaAxis(cb.Component.V, maximumVerticalSampling, height, out ChromaAxis cbRows) ||
            !TryCreateChromaAxis(cr.Component.V, maximumVerticalSampling, height, out ChromaAxis crRows)) {
            return false;
        }

        var job = new YccHighQualityJob(output, width, luma, cb, cr, cbColumns, crColumns, cbRows, crRows, outputRgba);
        if ((long)width * height >= 262_144 && height >= 16 && Environment.ProcessorCount > 1) {
            Parallel.For(0, height, new ParallelOptions {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = Math.Min(8, Environment.ProcessorCount)
            }, job.ComposeRow);
        } else {
            for (int y = 0; y < height; y++) {
                cancellationToken.ThrowIfCancellationRequested();
                job.ComposeRow(y);
            }
        }
        return true;
    }

    private sealed class YccHighQualityJob {
        private readonly byte[] _output;
        private readonly int _width;
        private readonly BaselineComponentState _luma;
        private readonly BaselineComponentState _cb;
        private readonly BaselineComponentState _cr;
        private readonly ChromaAxis _cbColumns;
        private readonly ChromaAxis _crColumns;
        private readonly ChromaAxis _cbRows;
        private readonly ChromaAxis _crRows;
        private readonly bool _outputRgba;

        public YccHighQualityJob(byte[] output, int width,
            BaselineComponentState luma, BaselineComponentState cb, BaselineComponentState cr,
            ChromaAxis cbColumns, ChromaAxis crColumns, ChromaAxis cbRows, ChromaAxis crRows, bool outputRgba) {
            _output = output;
            _width = width;
            _luma = luma;
            _cb = cb;
            _cr = cr;
            _cbColumns = cbColumns;
            _crColumns = crColumns;
            _cbRows = cbRows;
            _crRows = crRows;
            _outputRgba = outputRgba;
        }

        public void ComposeRow(int y) =>
            ComposeYccRow(_output, y * _width * (_outputRgba ? 4 : 3), _width, _outputRgba,
                _luma.Buffer, y * _luma.Stride,
                _cb.Buffer, _cbRows.First[y] * _cb.Stride, _cb.Buffer, _cbRows.Second[y] * _cb.Stride, _cbRows.Weight[y], _cbColumns,
                _cr.Buffer, _crRows.First[y] * _cr.Stride, _cr.Buffer, _crRows.Second[y] * _cr.Stride, _crRows.Weight[y], _crColumns);
    }

    // Converts one output row. Chroma rows may live in different buffers (adjacent bands).
    private static void ComposeYccRow(byte[] output, int target, int width, bool outputRgba,
        byte[] luma, int lumaRow,
        byte[] cbTopSamples, int cbTop, byte[] cbBottomSamples, int cbBottom, int cbRowWeight, ChromaAxis cbColumns,
        byte[] crTopSamples, int crTop, byte[] crBottomSamples, int crBottom, int crRowWeight, ChromaAxis crColumns) {
        int[] cbFirst = cbColumns.First;
        int[] cbSecond = cbColumns.Second;
        int[] cbWeight = cbColumns.Weight;
        int[] crFirst = crColumns.First;
        int[] crSecond = crColumns.Second;
        int[] crWeight = crColumns.Weight;
        double[] redTerms = YccHighQualityTables.Red;
        double[] greenCbTerms = YccHighQualityTables.GreenCb;
        double[] greenCrTerms = YccHighQualityTables.GreenCr;
        double[] blueTerms = YccHighQualityTables.Blue;
        int pixelBytes = outputRgba ? 4 : 3;

        for (int x = 0; x < width; x++, target += pixelBytes) {
            int cb16 = Interpolate16(cbTopSamples, cbTop, cbBottomSamples, cbBottom, cbFirst[x], cbSecond[x], cbWeight[x], cbRowWeight);
            int cr16 = Interpolate16(crTopSamples, crTop, crBottomSamples, crBottom, crFirst[x], crSecond[x], crWeight[x], crRowWeight);
            int lumaSample = luma[lumaRow + x];
            if (((cb16 | cr16) & 15) == 0) {
                int cbSample = cb16 >> 4;
                int crSample = cr16 >> 4;
                output[target] = ClampToByte(lumaSample + CrToR[crSample]);
                output[target + 1] = ClampToByte(lumaSample + ((CbToG[cbSample] + CrToG[crSample]) >> 16));
                output[target + 2] = ClampToByte(lumaSample + CbToB[cbSample]);
            } else {
                double lumaValue = lumaSample;
                output[target] = RoundColorSample(lumaValue + redTerms[cr16]);
                output[target + 1] = RoundColorSample(lumaValue - greenCbTerms[cb16] - greenCrTerms[cr16]);
                output[target + 2] = RoundColorSample(lumaValue + blueTerms[cb16]);
            }
            if (pixelBytes == 4) output[target + 3] = byte.MaxValue;
        }
    }

    // Bilinear chroma in sixteenths: quarter weights in each direction.
    private static int Interpolate16(byte[] topSamples, int topRow, byte[] bottomSamples, int bottomRow,
        int firstColumn, int secondColumn, int columnWeight, int rowWeight) {
        int topLeft = topSamples[topRow + firstColumn];
        int top = (topLeft << 2) + (topSamples[topRow + secondColumn] - topLeft) * columnWeight;
        int bottomLeft = bottomSamples[bottomRow + firstColumn];
        int bottom = (bottomLeft << 2) + (bottomSamples[bottomRow + secondColumn] - bottomLeft) * columnWeight;
        return (top << 2) + (bottom - top) * rowWeight;
    }

    // Same as ProjectColorSample for the bounded values produced here: truncation equals
    // Math.Floor for non-negative sums, and every negative sum clamps to zero either way.
    private static byte RoundColorSample(double value) {
        int rounded = (int)(value + .5D);
        if (rounded <= 0) return 0;
        if (rounded >= 255) return 255;
        return (byte)rounded;
    }

    private sealed class ChromaAxis {
        public int[] First = Array.Empty<int>();
        public int[] Second = Array.Empty<int>();
        // Weight of the second sample in quarters.
        public int[] Weight = Array.Empty<int>();
    }

    // Mirrors SampleComponentBilinear's coordinate mapping and edge replication for one axis.
    private static bool TryCreateChromaAxis(int sampling, int maximumSampling, int imageLength, out ChromaAxis axis) {
        axis = new ChromaAxis();
        if (sampling != maximumSampling && sampling * 2 != maximumSampling) return false;
        var first = new int[imageLength];
        var second = new int[imageLength];
        var weight = new int[imageLength];
        if (sampling == maximumSampling) {
            for (int i = 0; i < imageLength; i++) {
                first[i] = i;
                second[i] = i;
            }
        } else {
            int componentLength = (imageLength * sampling + maximumSampling - 1) / maximumSampling;
            for (int i = 0; i < imageLength; i++) {
                double position = (i + 0.5) * sampling / maximumSampling - 0.5;
                int lower = (int)Math.Floor(position);
                int upper = lower + 1;
                if (lower < 0) lower = 0;
                if (upper >= componentLength) upper = componentLength - 1;
                first[i] = lower;
                second[i] = upper;
                // Equal samples make the fraction irrelevant, including the negative
                // fraction left by clamping the leading edge.
                weight[i] = lower == upper ? 0 : (int)((position - lower) * 4);
            }
        }
        axis.First = first;
        axis.Second = second;
        axis.Weight = weight;
        return true;
    }

    private static class YccHighQualityTables {
        private const int Length = 255 * 16 + 1;
        public static readonly double[] Red = Create(1.402D);
        public static readonly double[] GreenCb = Create(.344136286D);
        public static readonly double[] GreenCr = Create(.714136286D);
        public static readonly double[] Blue = Create(1.772D);

        private static double[] Create(double factor) {
            var table = new double[Length];
            for (int i = 0; i < table.Length; i++) {
                double centered = i / 16D;
                centered -= 128;
                table[i] = factor * centered;
            }
            return table;
        }
    }
}
