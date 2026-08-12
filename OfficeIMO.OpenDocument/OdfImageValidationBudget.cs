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
        // Reserve the cheap aggregate limits before parsing any attacker-controlled
        // payload. Failed identification still consumes the attempt and encoded-byte
        // allowance, so malformed images cannot bypass the conversion-wide budget.
        if (_images >= MaximumImages ||
            _encodedBytes >= MaximumEncodedBytes ||
            _pixels >= MaximumPixels) {
            return false;
        }

        _images++;
        if (_encodedBytes > MaximumEncodedBytes - bytes.LongLength) {
            _encodedBytes = MaximumEncodedBytes;
            return false;
        }
        _encodedBytes += bytes.LongLength;

        if (!OfficeImageReader.TryIdentifyByContent(bytes, fileName, out OfficeImageInfo info)) return false;
        if (info.Format == OfficeImageFormat.Svg) return true;

        long pixels = (long)info.Width * info.Height;
        if (pixels <= 0) return false;
        if (_pixels > MaximumPixels - pixels) {
            _pixels = MaximumPixels;
            return false;
        }

        _pixels += pixels;
        return true;
    }
}
