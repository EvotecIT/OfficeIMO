#if NET10_0_OR_GREATER
using System.Buffers;
using System.Runtime.InteropServices;
using OneImage;
using OneImageImage = OneImage.Image;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>
/// OneImage lanes for the image comparisons. OneImage ships net10.0 only, so these adapters compile only for that
/// target. Every lane uses the public API and the library's documented defaults unless a setting has to match the
/// work the other engines do.
/// </summary>
internal static class OneImageComparisonAdapters {
    // The other decoders return the stored pixel layout; OneImage applies EXIF orientation unless asked not to.
    private static readonly LoadOptions StoredLayout = new() { AutoOrient = false };

    // Deflate level 6 is OneImage's default and the zlib level SkiaSharp and OfficeIMO's optimal setting use.
    private static readonly SaveOptions Png = new() { EncoderOptions = new PngEncoderOptions { CompressionLevel = 6 } };

    // OfficeIMO writes LZW with the horizontal predictor; OneImage uses the same TIFF 6.0 predictor.
    private static readonly SaveOptions TiffLzw = new() {
        EncoderOptions = new TiffEncoderOptions { Compression = TiffCompression.Lzw, UsePredictor = true }
    };

    // Effort 4 is OneImage's default and libwebp's default method, which Magick.NET uses.
    private static readonly SaveOptions WebpLossless = new() {
        EncoderOptions = new WebPEncoderOptions { Lossless = true, Effort = 4 }
    };

    /// <summary>Decodes into a caller-sized RGBA buffer, the same contract as the SkiaSharp lane.</summary>
    internal static byte[] DecodeRgba(byte[] encoded) {
        OneImageImage image = OneImageImage.Load(encoded, StoredLayout);
        var rgba = new byte[checked(image.Width * image.Height * 4)];
        image.CopyPixelsTo(MemoryMarshal.Cast<byte, Rgba32>(rgba.AsSpan()), new ImageRect(0, 0, image.Width, image.Height));
        return rgba;
    }

    internal static OneImageImage CreateImage(byte[] rgba, int width, int height) =>
        OneImageImage.FromPixels<Rgba32>(MemoryMarshal.Cast<byte, Rgba32>(rgba.AsSpan()), width, height);

    internal static byte[] EncodePng(OneImageImage image) => Save(image, PngFormat.Instance, Png);

    internal static byte[] EncodeTiffLzw(OneImageImage image) => Save(image, TiffFormat.Instance, TiffLzw);

    internal static byte[] EncodeWebpLossless(OneImageImage image) => Save(image, WebPFormat.Instance, WebpLossless);

    /// <summary>
    /// Resizes with OneImage's span-to-span resizer, which its documentation names as the path for repeated resizes.
    /// The triangle filter is OneImage's bilinear; filtering runs on the stored sRGB values like the other lanes.
    /// </summary>
    internal static byte[] ResizeRgba(PixelResizer resizer, byte[] rgba, int sourceWidth, int sourceHeight, int width, int height) {
        var resized = new byte[checked(width * height * 4)];
        resizer.Resize(
            MemoryMarshal.Cast<byte, Rgba32>(rgba.AsSpan()), sourceWidth, sourceHeight,
            MemoryMarshal.Cast<byte, Rgba32>(resized.AsSpan()), width, height,
            ResamplingFilter.Bilinear, linear: false);
        return resized;
    }

    private static byte[] Save(OneImageImage image, IImageFormat format, SaveOptions options) {
        var writer = new ArrayBufferWriter<byte>();
        image.Save(writer, format, options);
        return writer.WrittenSpan.ToArray();
    }
}
#endif
