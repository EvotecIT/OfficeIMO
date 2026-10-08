using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterResampler {
    /// <summary>Plans the managed working-set bytes required to resize an image without allocating its pixel buffers.</summary>
    /// <remarks>The result includes source and destination RGBA storage, resampling scratch, contribution tables and fixed working overhead. Identical dimensions use the independent-copy path. Dimensions or sampling work exceeding the managed image limits are rejected.</remarks>
    /// <param name="sourceWidth">Source width in pixels.</param>
    /// <param name="sourceHeight">Source height in pixels.</param>
    /// <param name="width">Destination width in pixels.</param>
    /// <param name="height">Destination height in pixels.</param>
    /// <param name="mode">The sampling kernel used by the resize operation.</param>
    /// <param name="cancellationToken">Observes cancellation during contribution planning.</param>
    /// <returns>The planned managed working-set bytes for this operation.</returns>
    public static long GetResizeWorkingSetBytes(int sourceWidth, int sourceHeight, int width, int height,
        OfficeRasterResamplingMode mode = OfficeRasterResamplingMode.Bilinear, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterGuards.EnsureOutputPixels(sourceWidth, sourceHeight, "Source dimensions exceed the managed image limit.");
        OfficeRasterGuards.EnsureOutputPixels(width, height, "Raster resize dimensions exceed the managed image limit.");
        if (mode < OfficeRasterResamplingMode.NearestNeighbor || mode > OfficeRasterResamplingMode.Welch) {
            throw new ArgumentOutOfRangeException(nameof(mode));
        }
        bool simple = (sourceWidth == width && sourceHeight == height) ||
            mode == OfficeRasterResamplingMode.NearestNeighbor || mode == OfficeRasterResamplingMode.Bilinear;
        long workingSetBytes;
        bool withinLimit = simple
            ? TryGetSimpleWorkingSetBytes(sourceWidth, sourceHeight, width, height, retainedManagedBytes: 0, out workingSetBytes)
            : TryMeasureSeparableWorkingSet(sourceWidth, sourceHeight, width, height, mode, retainedManagedBytes: 0,
                out _, out _, out _, out workingSetBytes, cancellationToken);
        if (!withinLimit) {
            throw new ArgumentException("Raster resampling working set exceeds the managed image limit.");
        }
        cancellationToken.ThrowIfCancellationRequested();
        return workingSetBytes;
    }
}
