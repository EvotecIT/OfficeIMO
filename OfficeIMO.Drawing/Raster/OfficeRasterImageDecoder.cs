namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free decoder for raster image bytes that can be painted by <see cref="OfficeRasterCanvas"/>.
/// </summary>
public static class OfficeRasterImageDecoder {
    /// <summary>
    /// Human-readable summary of raster formats currently decoded by the managed renderer.
    /// </summary>
    public const string SupportedFormatDescription = "PNG, JPEG, baseline RGB/RGBA TIFF, uncompressed BMP, explicitly selected GIF frames, and OfficeIMO literal-lossless WebP image bytes";

    /// <summary>
    /// Attempts to decode image bytes into an RGBA raster buffer supported by dependency-free export.
    /// </summary>
    public static bool TryDecode(byte[]? bytes, out OfficeRasterImage? image) =>
        TryDecode(bytes, options: null, out image, out _);

    /// <summary>
    /// Attempts to decode image bytes using explicit frame and animation-loss policy.
    /// </summary>
    public static bool TryDecode(
        byte[]? bytes,
        OfficeRasterDecodeOptions? options,
        out OfficeRasterImage? image,
        out OfficeRasterDecodeInfo info) {
        image = null;
        var effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        OfficeImageFormat format = IdentifyFormat(bytes);
        if (bytes == null || bytes.Length == 0) {
            info = new OfficeRasterDecodeInfo(format, 0, effective.FrameIndex, succeeded: false, diagnostic: "Raster image bytes are empty.");
            return false;
        }

        if (format == OfficeImageFormat.Gif) {
            bool decoded = OfficeGifReader.TryDecodeFrame(bytes, effective.FrameIndex, out image, out int frameCount);
            if (effective.AnimationPolicy == OfficeRasterAnimationPolicy.RejectAnimated && frameCount > 1) {
                image = null;
                info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, succeeded: false, diagnostic: "Animated GIF input was rejected by the configured animation policy.");
                return false;
            }

            string? diagnostic = decoded && frameCount > 1
                ? "The selected GIF frame was decoded; remaining animation frames were not retained in the static raster result."
                : decoded ? null : "The requested GIF frame could not be decoded.";
            info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, decoded, diagnostic);
            return decoded;
        }

        int webpFrameCount = format == OfficeImageFormat.Webp ? CountWebpAnimationFrames(bytes) : 1;
        if (webpFrameCount > 1) {
            info = new OfficeRasterDecodeInfo(format, webpFrameCount, effective.FrameIndex, succeeded: false,
                diagnostic: effective.AnimationPolicy == OfficeRasterAnimationPolicy.RejectAnimated
                    ? "Animated WebP input was rejected by the configured animation policy."
                    : "Animated WebP frame decoding is outside the managed literal-lossless WebP subset.");
            return false;
        }

        if (effective.FrameIndex != 0) {
            info = new OfficeRasterDecodeInfo(format, 1, effective.FrameIndex, succeeded: false, diagnostic: "The selected raster format exposes only frame zero through the managed decoder.");
            return false;
        }

        bool success =
            OfficePngReader.TryDecode(bytes, out image) ||
            OfficeJpegCodec.TryDecode(bytes, out image) ||
            OfficeTiffCodec.TryDecode(bytes, out image) ||
            OfficeBmpReader.TryDecode(bytes, out image) ||
            OfficeWebpCodec.TryDecode(bytes, out image);
        info = new OfficeRasterDecodeInfo(format, success ? 1 : 0, effective.FrameIndex, success,
            success ? null : "Raster bytes are not supported by the managed decoder subset.");
        return success;
    }

    private static OfficeImageFormat IdentifyFormat(byte[]? bytes) =>
        bytes != null && OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo identified)
            ? identified.Format
            : OfficeImageFormat.Unknown;

    private static int CountWebpAnimationFrames(byte[] bytes) {
        if (bytes.Length < 12 ||
            bytes[0] != (byte)'R' || bytes[1] != (byte)'I' || bytes[2] != (byte)'F' || bytes[3] != (byte)'F' ||
            bytes[8] != (byte)'W' || bytes[9] != (byte)'E' || bytes[10] != (byte)'B' || bytes[11] != (byte)'P') {
            return 1;
        }

        int count = 0;
        int offset = 12;
        while (offset + 8 <= bytes.Length) {
            int length = bytes[offset + 4] |
                (bytes[offset + 5] << 8) |
                (bytes[offset + 6] << 16) |
                (bytes[offset + 7] << 24);
            if (length < 0 || (long)offset + 8L + length > bytes.Length) break;
            if (bytes[offset] == (byte)'A' && bytes[offset + 1] == (byte)'N' &&
                bytes[offset + 2] == (byte)'M' && bytes[offset + 3] == (byte)'F') count++;
            offset += 8 + length + (length & 1);
        }
        return count > 0 ? count : 1;
    }
}
