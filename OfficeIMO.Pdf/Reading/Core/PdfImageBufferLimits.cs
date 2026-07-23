namespace OfficeIMO.Pdf;

internal static class PdfImageBufferLimits {
    internal static bool TryGetScanlineBufferSize(
        int width,
        int height,
        int channels,
        out int pixelCount,
        out int bufferSize) {
        pixelCount = 0;
        bufferSize = 0;
        if (width <= 0 || height <= 0 || channels <= 0) return false;

        long pixels = (long)width * height;
        long rowBytes = 1L + (long)width * channels;
        long totalBytes = rowBytes * height;
        if (pixels > int.MaxValue || rowBytes > int.MaxValue || totalBytes > int.MaxValue ||
            totalBytes > PdfReadLimits.DefaultMaxDecodedStreamBytes) return false;

        pixelCount = (int)pixels;
        bufferSize = (int)totalBytes;
        return true;
    }
}
