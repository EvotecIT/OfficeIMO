using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Bounds aggregate image-validation work during one document conversion.</summary>
internal sealed class OdfImageValidationBudget {
    private const int MaximumImages = 256;
    private const long MaximumEncodedBytes = 128L * 1024L * 1024L;
    private const long MaximumPixels = 100_000_000L;
    private int _images;
    private long _encodedBytes;
    private long _pixels;

    internal bool TryConsume(byte[] bytes, string? fileName) {
        if (!OfficeImageReader.TryIdentifyByContent(bytes, fileName, out OfficeImageInfo info)) return false;
        long pixels = (long)info.Width * info.Height;
        if (pixels <= 0 ||
            _images >= MaximumImages ||
            _encodedBytes > MaximumEncodedBytes - bytes.LongLength ||
            _pixels > MaximumPixels - pixels) {
            return false;
        }

        _images++;
        _encodedBytes += bytes.LongLength;
        _pixels += pixels;
        return true;
    }
}
