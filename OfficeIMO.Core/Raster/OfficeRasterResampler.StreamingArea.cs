namespace OfficeIMO.Drawing;

public static partial class OfficeRasterResampler {
    // Three rows of four double accumulators per output column; each row costs eight RGBA pixels per column.
    private const long StreamingAreaAccumulatorPixelsPerColumn = 24L;

    /// <summary>Working pixels for <see cref="MinifyByStreamingArea"/>: its output plus three accumulator rows.</summary>
    internal static long GetStreamingAreaPixelCost(int width, int height) =>
        checked((long)width * height + width * StreamingAreaAccumulatorPixelsPerColumn);

    /// <summary>
    /// Measures the managed bytes held while <see cref="MinifyByStreamingArea"/> runs: the source, the output,
    /// the accumulator rows, the per-axis coverage tables and any retained caller allocation such as the canvas.
    /// </summary>
    internal static bool TryGetStreamingAreaWorkingSetBytes(
        int sourceWidth,
        int sourceHeight,
        int width,
        int height,
        long retainedManagedBytes,
        out long workingSetBytes) {
        workingSetBytes = 0L;
        if (sourceWidth <= 0 || sourceHeight <= 0 || width <= 0 || height <= 0 ||
            width > sourceWidth || height > sourceHeight || retainedManagedBytes < 0L) return false;
        try {
            long sourceBytes = checked((long)sourceWidth * sourceHeight * 4L);
            long outputBytes = checked((long)width * height * 4L);
            long accumulatorBytes = checked(3L * width * 4L * sizeof(double));
            // Per source column and row: the first covered output index and its weight.
            long coverageBytes = checked(((long)sourceWidth + sourceHeight) * (sizeof(int) + sizeof(double)));
            workingSetBytes = checked(sourceBytes + outputBytes + accumulatorBytes + coverageBytes +
                retainedManagedBytes + 64L * 1024L);
            return workingSetBytes <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    /// <summary>
    /// Area-averages a source into smaller or equal dimensions in premultiplied encoded-sRGB space,
    /// one source row at a time. Every output pixel covers exactly its share of the source, including
    /// at the right and bottom edges, and no full-height float scratch is allocated.
    /// </summary>
    internal static OfficeRasterImage MinifyByStreamingArea(
        OfficeRasterImage source, int width, int height, CancellationToken cancellationToken) {
        if (width < 1 || width > source.Width) throw new ArgumentOutOfRangeException(nameof(width));
        if (height < 1 || height > source.Height) throw new ArgumentOutOfRangeException(nameof(height));
        CreateAreaCoverage(source.Width, width, out int[] firstColumn, out double[] firstColumnWeight);
        CreateAreaCoverage(source.Height, height, out int[] firstRow, out double[] firstRowWeight);
        byte[] input = source.PixelBuffer;
        byte[] output = new byte[checked(width * height * 4)];
        int rowLength = checked(width * 4);
        double[] sourceRow = new double[rowLength];
        double[] current = new double[rowLength];
        double[] next = new double[rowLength];
        double outputArea = (double)source.Width / width * ((double)source.Height / height);
        int currentRow = 0;
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            Array.Clear(sourceRow, 0, sourceRow.Length);
            int row = y * source.Width * 4;
            for (int x = 0; x < source.Width; x++) {
                int pixel = row + x * 4;
                double alpha = input[pixel + 3];
                if (alpha <= 0D) continue;
                double red = input[pixel] * alpha;
                double green = input[pixel + 1] * alpha;
                double blue = input[pixel + 2] * alpha;
                int column = firstColumn[x] * 4;
                double weight = firstColumnWeight[x];
                sourceRow[column] += red * weight;
                sourceRow[column + 1] += green * weight;
                sourceRow[column + 2] += blue * weight;
                sourceRow[column + 3] += alpha * weight;
                if (weight < 1D) {
                    weight = 1D - weight;
                    sourceRow[column + 4] += red * weight;
                    sourceRow[column + 5] += green * weight;
                    sourceRow[column + 6] += blue * weight;
                    sourceRow[column + 7] += alpha * weight;
                }
            }
            while (currentRow < firstRow[y]) {
                WriteAreaRow(current, output, currentRow * rowLength, outputArea);
                (current, next) = (next, current);
                Array.Clear(next, 0, next.Length);
                currentRow++;
            }
            double rowWeight = firstRowWeight[y];
            for (int index = 0; index < rowLength; index++) current[index] += sourceRow[index] * rowWeight;
            if (rowWeight < 1D) {
                rowWeight = 1D - rowWeight;
                for (int index = 0; index < rowLength; index++) next[index] += sourceRow[index] * rowWeight;
            }
        }
        while (currentRow < height) {
            WriteAreaRow(current, output, currentRow * rowLength, outputArea);
            (current, next) = (next, current);
            Array.Clear(next, 0, next.Length);
            currentRow++;
        }
        return OfficeRasterImage.FromOwnedRgba32(width, height, output);
    }

    // Source pixel i spans [i, i + 1) and output pixel j spans [j * step, (j + 1) * step) with step >= 1,
    // so a source pixel lies inside one output pixel or straddles exactly one boundary. Boundaries are
    // compared in integers: output j ends at source coordinate (j + 1) * sourceLength / length.
    private static void CreateAreaCoverage(int sourceLength, int length, out int[] first, out double[] firstWeight) {
        first = new int[sourceLength];
        firstWeight = new double[sourceLength];
        for (int index = 0; index < sourceLength; index++) {
            int start = (int)((long)index * length / sourceLength);
            long boundary = (long)(start + 1) * sourceLength;
            first[index] = start;
            firstWeight[index] = start + 1 < length && (long)(index + 1) * length > boundary
                ? (boundary - (long)index * length) / (double)length
                : 1D;
        }
    }

    private static void WriteAreaRow(double[] sums, byte[] output, int target, double outputArea) {
        for (int index = 0; index < sums.Length; index += 4) {
            double alphaSum = sums[index + 3];
            if (alphaSum <= 0D) continue;
            output[target + index] = ToByte(sums[index] / alphaSum);
            output[target + index + 1] = ToByte(sums[index + 1] / alphaSum);
            output[target + index + 2] = ToByte(sums[index + 2] / alphaSum);
            output[target + index + 3] = ToByte(alphaSum / outputArea);
        }
    }
}
