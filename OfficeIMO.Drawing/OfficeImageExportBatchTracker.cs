using System;

namespace OfficeIMO.Drawing;

internal sealed class OfficeImageExportBatchTracker {
    private readonly OfficeImageExportOptions _options;

    internal OfficeImageExportBatchTracker(OfficeImageExportOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    internal int Count { get; private set; }

    internal long TotalRasterPixels { get; private set; }

    internal long TotalEncodedBytes { get; private set; }

    internal void Add(OfficeImageExportResult result) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        int nextCount = checked(Count + 1);
        EnsureWithin(nameof(OfficeImageExportOptions.MaximumOutputCount), nextCount, _options.MaximumOutputCount);

        long nextPixels = TotalRasterPixels;
        if (result.Format.IsRaster()) {
            long pixels = checked((long)result.Width * result.Height);
            nextPixels = checked(TotalRasterPixels + pixels);
            EnsureWithin(nameof(OfficeImageExportOptions.MaximumTotalRasterPixels), nextPixels, _options.MaximumTotalRasterPixels);
        }

        long nextBytes = checked(TotalEncodedBytes + result.EncodedLength);
        EnsureWithin(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), nextBytes, _options.MaximumTotalEncodedBytes);

        Count = nextCount;
        TotalRasterPixels = nextPixels;
        TotalEncodedBytes = nextBytes;
    }

    private static void EnsureWithin(string name, long actual, long maximum) {
        if (actual > maximum) throw new OfficeImageExportBatchLimitException(name, actual, maximum);
    }
}
