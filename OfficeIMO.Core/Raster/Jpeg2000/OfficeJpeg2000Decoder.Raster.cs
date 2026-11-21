namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    /// <summary>
    /// Decodes a JP2 file or codestream to RGBA for the raster pipeline: gray, RGB and sYCC with optional alpha,
    /// palettes, and CMYK through a device conversion. ICC profiles are not applied; their samples are used as
    /// gray or RGB by channel count.
    /// </summary>
    internal static bool TryDecodeRaster(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterImage? image) {
        image = null;
        if (options.FrameIndex != 0) return false;
        var decodeOptions = new OfficeJpeg2000DecodeOptions {
            MaximumPixels = options.MaximumDecodedPixels,
            MaximumWorkingBytes = Math.Max(0L, OfficeRasterGuards.MaximumDecodedBytes - options.RetainedManagedBytes),
            CancellationToken = options.CancellationToken
        };
        if (!TryDecode(bytes, decodeOptions, out OfficeJpeg2000Image? decoded) || decoded == null) return false;
        int colors = decoded.ColorChannels;
        bool cmyk = colors == 4 && decoded.ColorSpace is OfficeJpeg2000ColorSpace.Cmyk or OfficeJpeg2000ColorSpace.Unspecified or OfficeJpeg2000ColorSpace.Icc;
        if (colors != 1 && colors != 3 && !cmyk) return false;
        long pixels = (long)decoded.Width * decoded.Height;
        if (pixels * 4 > int.MaxValue) return false;
        var rgba = new byte[pixels * 4];
        byte[] samples = decoded.Samples;
        int channels = decoded.Channels;
        for (long pixel = 0, source = 0, target = 0; pixel < pixels; pixel++, source += channels, target += 4) {
            if ((pixel & 0xFFFF) == 0) options.CancellationToken.ThrowIfCancellationRequested();
            byte alpha = decoded.HasAlpha ? samples[source + colors] : (byte)255;
            int red, green, blue;
            if (colors == 1) {
                red = green = blue = samples[source];
            } else if (colors == 3) {
                red = samples[source];
                green = samples[source + 1];
                blue = samples[source + 2];
            } else {
                int black = 255 - samples[source + 3];
                red = (255 - samples[source]) * black / 255;
                green = (255 - samples[source + 1]) * black / 255;
                blue = (255 - samples[source + 2]) * black / 255;
            }
            if (decoded.PremultipliedAlpha && alpha != 255) {
                red = alpha == 0 ? 0 : Math.Min(255, red * 255 / alpha);
                green = alpha == 0 ? 0 : Math.Min(255, green * 255 / alpha);
                blue = alpha == 0 ? 0 : Math.Min(255, blue * 255 / alpha);
            }
            rgba[target] = (byte)red;
            rgba[target + 1] = (byte)green;
            rgba[target + 2] = (byte)blue;
            rgba[target + 3] = alpha;
        }
        image = OfficeRasterImage.FromOwnedRgba32(decoded.Width, decoded.Height, rgba);
        return true;
    }
}
