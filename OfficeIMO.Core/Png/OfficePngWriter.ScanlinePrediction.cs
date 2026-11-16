using System.IO.Compression;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    private const int PredictionSampleBandRows = 8;
    private const int PredictionSampleBytes = 256 * 1024;

    /// <summary>
    /// Replaces unfiltered scanlines (each prefixed by filter type 0) with adaptively filtered ones when that
    /// compresses smaller. Photos, gradients and logos shrink with row prediction, while rendered pages and
    /// screenshots usually compress best unfiltered because prediction breaks up their repeated runs.
    /// </summary>
    /// <returns><see langword="true"/> when the rows were filtered.</returns>
    internal static bool PredictScanlinesWhenSmaller(byte[] scanlines, int rowBytes, int height, int bytesPerPixel,
        CancellationToken token = default) {
        if (scanlines == null) throw new ArgumentNullException(nameof(scanlines));
        if (bytesPerPixel <= 0) throw new ArgumentOutOfRangeException(nameof(bytesPerPixel));
        if (rowBytes <= 0 || height <= 0) return false;
        if ((long)(rowBytes + 1L) * height > scanlines.LongLength)
            throw new ArgumentException("The scanline buffer is shorter than its dimensions.", nameof(scanlines));
        var workspace = new PngColorWorkspace(rowBytes);
        if (!PredictionCompressesSmaller(scanlines, rowBytes, height, bytesPerPixel, workspace, token)) return false;
        FilterScanlines(scanlines, rowBytes, height, bytesPerPixel, workspace, token);
        return true;
    }

    // Deflates evenly spaced bands of rows quickly, unfiltered and filtered, instead of compressing the image twice.
    private static bool PredictionCompressesSmaller(byte[] scanlines, int rowBytes, int height, int bytesPerPixel,
        PngColorWorkspace workspace, CancellationToken token) {
        int rowLength = rowBytes + 1;
        // Very wide rows shorten the bands so the sample stays near its byte budget; one row is the minimum.
        int bandRows = Math.Max(1, Math.Min(Math.Min(height, PredictionSampleBandRows), PredictionSampleBytes / rowLength));
        int bands = Math.Max(1, Math.Min(height / bandRows, PredictionSampleBytes / Math.Max(1, rowLength * bandRows)));
        var raw = new byte[checked(bands * bandRows * rowLength)];
        var filtered = new byte[raw.Length];
        int target = 0;
        for (int band = 0; band < bands; band++) {
            token.ThrowIfCancellationRequested();
            int start = bands == 1 ? 0 : (int)((long)band * (height - bandRows) / (bands - 1));
            if (start > 0) Buffer.BlockCopy(scanlines, (start - 1) * rowLength + 1, workspace.Previous, 0, rowBytes);
            else Array.Clear(workspace.Previous, 0, rowBytes);
            for (int y = start; y < start + bandRows; y++, target += rowLength) {
                Buffer.BlockCopy(scanlines, y * rowLength, raw, target, rowLength);
                FilterRow(scanlines, y * rowLength, rowBytes, bytesPerPixel, workspace, token);
                Buffer.BlockCopy(workspace.Selected, 0, filtered, target, rowLength);
            }
        }
        return DeflatedLength(filtered, token) < DeflatedLength(raw, token);
    }

    private static void FilterScanlines(byte[] scanlines, int rowBytes, int height, int bytesPerPixel,
        PngColorWorkspace workspace, CancellationToken token) {
        int rowLength = rowBytes + 1;
        Array.Clear(workspace.Previous, 0, rowBytes);
        for (int y = 0, offset = 0; y < height; y++, offset += rowLength) {
            token.ThrowIfCancellationRequested();
            FilterRow(scanlines, offset, rowBytes, bytesPerPixel, workspace, token);
            Buffer.BlockCopy(workspace.Selected, 0, scanlines, offset, rowLength);
        }
    }

    // Filters the row at offset against workspace.Previous, then makes the unfiltered row the next one's predecessor.
    private static void FilterRow(byte[] scanlines, int offset, int rowBytes, int bytesPerPixel, PngColorWorkspace workspace,
        CancellationToken token) {
        Buffer.BlockCopy(scanlines, offset + 1, workspace.Current, 0, rowBytes);
        // Sampled scoring keeps prediction close to the cost of the deflate it shrinks.
        SelectRowFilter(workspace.Current, workspace.Previous, workspace.Selected, rowBytes, bytesPerPixel,
            adaptive: true, sampledFilters: true, token, observer: null);
        var swap = workspace.Previous; workspace.Previous = workspace.Current; workspace.Current = swap;
    }

    private static long DeflatedLength(byte[] data, CancellationToken token) {
        using var probe = new PngSizeProbeStream();
        using (var deflate = new DeflateStream(probe, CompressionLevel.Fastest, leaveOpen: true)) {
            for (int position = 0; position < data.Length; position += 65536) {
                token.ThrowIfCancellationRequested();
                deflate.Write(data, position, Math.Min(65536, data.Length - position));
            }
        }
        return probe.Length;
    }
}
