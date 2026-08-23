using System;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free decoder for raster image bytes that can be painted by <see cref="OfficeRasterCanvas"/>.
/// </summary>
public static class OfficeRasterImageDecoder {
    /// <summary>
    /// Human-readable summary of raster formats currently decoded by the managed renderer.
    /// </summary>
    public const string SupportedFormatDescription = "PNG and APNG frames, JPEG, bounded classic TIFF pages, uncompressed BMP, explicitly selected GIF frames, and lossless VP8L WebP image bytes";

    /// <summary>
    /// Attempts to decode image bytes into an RGBA raster buffer supported by dependency-free export.
    /// </summary>
    public static bool TryDecode(byte[]? bytes, out OfficeRasterImage? image) =>
        TryDecode(bytes, options: null, out image, out _);

    /// <summary>
    /// Attempts to decode a bounded raster stream and leaves a seekable stream at its original position.
    /// </summary>
    public static bool TryDecode(Stream stream, out OfficeRasterImage? image) =>
        TryDecode(stream, options: null, out image, out _);

    /// <summary>
    /// Attempts to decode a bounded raster stream using explicit selection, loss, resource, and cancellation policy.
    /// </summary>
    public static bool TryDecode(
        Stream stream,
        OfficeRasterDecodeOptions? options,
        out OfficeRasterImage? image,
        out OfficeRasterDecodeInfo info) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        OfficeRasterDecodeOptions effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        long originalPosition = stream.CanSeek ? stream.Position : 0L;
        try {
            if (!OfficeBoundedStreamReader.TryRead(stream, effective.MaximumEncodedBytes, effective.CancellationToken, out byte[] bytes)) {
                image = null;
                info = new OfficeRasterDecodeInfo(OfficeImageFormat.Unknown, 0, effective.FrameIndex, false,
                    "The raster stream is empty or exceeds the configured encoded-size limit.");
                return false;
            }
            return TryDecode(bytes, effective, out image, out info);
        } finally {
            if (stream.CanSeek) stream.Position = originalPosition;
        }
    }

    internal static bool TryDecode(byte[]? bytes, long maximumRasterPixels, out OfficeRasterImage? image) {
        if (maximumRasterPixels <= 0L) throw new System.ArgumentOutOfRangeException(nameof(maximumRasterPixels));
        OfficeImageFormat format = IdentifyFormat(bytes);
        if (IsManagedRasterFormat(format) &&
            OfficeImageReader.TryIdentifyByContent(bytes, fileName: null, out OfficeImageInfo info) &&
            !IsWithinPixelLimit(info.Width, info.Height, maximumRasterPixels)) {
            image = null;
            return false;
        }

        if (!TryDecode(bytes, out image) || image == null) return false;
        if (IsWithinPixelLimit(image.Width, image.Height, maximumRasterPixels)) return true;
        image = null;
        return false;
    }

    internal static bool IsWithinPixelLimit(int width, int height, long maximumRasterPixels) =>
        width > 0 && height > 0 && width <= maximumRasterPixels && height <= maximumRasterPixels / width;

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
        effective.CancellationToken.ThrowIfCancellationRequested();
        OfficeImageFormat format = IdentifyFormat(bytes);
        if (bytes == null || bytes.Length == 0 || bytes.Length > effective.MaximumEncodedBytes) {
            info = new OfficeRasterDecodeInfo(format, 0, effective.FrameIndex, succeeded: false, diagnostic: "Raster image bytes are empty.");
            return false;
        }

        if (!OfficeRasterContainerInspector.TryInspect(bytes, effective, out OfficeRasterContainerInfo? container) || container == null) {
            info = new OfficeRasterDecodeInfo(format, 0, effective.FrameIndex, succeeded: false,
                diagnostic: "The raster container is malformed, unsupported, or outside the configured limits.");
            return false;
        }
        int frameCount = container.Count;
        if (effective.FrameIndex >= frameCount) {
            info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, false,
                "The requested frame or page index is outside the container.", container);
            return false;
        }
        if (frameCount > 1 && effective.FrameLossPolicy == OfficeRasterFrameLossPolicy.RejectMultipleFrames) {
            info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, false,
                container.IsMultiPage
                    ? "Multi-page TIFF input was rejected by the configured frame-loss policy."
                    : "Animated input was rejected by the configured frame-loss policy.",
                container);
            return false;
        }

        if (format == OfficeImageFormat.Gif) {
            bool decoded = OfficeGifReader.TryDecodeFrame(bytes, effective.FrameIndex, out image, out int decodedFrameCount);
            string? diagnostic = decoded && frameCount > 1
                ? "The selected GIF frame was decoded; remaining animation frames were not retained in the static raster result."
                : decoded ? null : "The requested GIF frame could not be decoded.";
            bool withinLimit = decoded && decodedFrameCount == frameCount && IsDecodedImageWithinLimit(image, effective.MaximumDecodedPixels);
            if (!withinLimit) image = null;
            info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, withinLimit,
                withinLimit ? diagnostic : "The requested GIF frame could not be decoded within the configured limits.", container);
            return withinLimit;
        }

        if (format == OfficeImageFormat.Png) {
            if (frameCount > 1) {
                bool decoded = OfficeApngDecoder.TryDecodeFrame(bytes, container, effective.FrameIndex,
                    effective.MaximumDecodedPixels, effective.CancellationToken, out image);
                if (!decoded) image = null;
                info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, decoded,
                    decoded
                        ? "The selected APNG frame was composed; remaining animation frames were not retained in the static raster result."
                        : "The selected APNG frame could not be decoded within the configured limits.", container);
                return decoded;
            }
        }

        if (format == OfficeImageFormat.Tiff) {
            bool decoded = OfficeTiffCodec.TryDecodePage(bytes, effective.FrameIndex, effective, out image);
            string? diagnostic = decoded && frameCount > 1
                ? "The selected TIFF page was decoded; remaining pages were not retained in the static raster result."
                : decoded ? null : "The requested TIFF page could not be decoded.";
            info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, decoded, diagnostic, container);
            return decoded;
        }

        if (format == OfficeImageFormat.Webp && container.IsAnimated) {
            info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, false,
                "Animated WebP pixel decoding remains an explicit caller-codec boundary.", container);
            return false;
        }

        effective.CancellationToken.ThrowIfCancellationRequested();
        bool success = format switch {
            OfficeImageFormat.Png => OfficePngReader.TryDecode(
                bytes, effective.CancellationToken, out image),
            OfficeImageFormat.Jpeg => OfficeJpegCodec.TryDecode(bytes, out image),
            OfficeImageFormat.Bmp => OfficeBmpReader.TryDecode(bytes, out image),
            OfficeImageFormat.Webp => OfficeWebpCodec.TryDecode(
                bytes, effective.CancellationToken, out image),
            _ => false
        };
        success = success && IsDecodedImageWithinLimit(image, effective.MaximumDecodedPixels);
        if (!success) image = null;
        info = new OfficeRasterDecodeInfo(format, frameCount, effective.FrameIndex, success,
            success ? null : "Raster bytes are not supported by the managed decoder subset or exceed configured limits.", container);
        return success;
    }

    private static bool IsDecodedImageWithinLimit(OfficeRasterImage? image, long maximumPixels) =>
        image != null && IsWithinPixelLimit(image.Width, image.Height, maximumPixels);

    private static OfficeImageFormat IdentifyFormat(byte[]? bytes) =>
        bytes != null && OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo identified)
            ? identified.Format
            : OfficeImageFormat.Unknown;

    private static bool IsManagedRasterFormat(OfficeImageFormat format) =>
        format == OfficeImageFormat.Png ||
        format == OfficeImageFormat.Jpeg ||
        format == OfficeImageFormat.Gif ||
        format == OfficeImageFormat.Bmp ||
        format == OfficeImageFormat.Tiff ||
        format == OfficeImageFormat.Webp;

}
