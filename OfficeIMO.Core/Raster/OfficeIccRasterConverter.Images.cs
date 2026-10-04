using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeIccRasterConverter {
    // The caller owns profile parsing/retention policy. Decode device channels before
    // color conversion: converting a CMYK decoder's approximate RGB output loses ink data.
    internal static bool TryDecodeToSrgb(byte[] encoded, OfficeIccColorProfile profile,
        OfficeRasterDecodeOptions options, out OfficeRasterImage? image) {
        image = null;
        options.Validate();
        options.CancellationToken.ThrowIfCancellationRequested();
        if (encoded.Length > options.MaximumEncodedBytes ||
            !OfficeImageReader.TryIdentifyByContent(encoded, null, options.CancellationToken, out var info) ||
            !OfficeRasterGuards.TryEnsurePixelCount(info.Width, info.Height, options.MaximumDecodedPixels, out int pixels)) return false;
        var effective = options.WithAdditionalRetainedManagedBytes(profile.RetainedByteCount);
        const OfficeIccRenderingIntent intent = OfficeIccRenderingIntent.RelativeColorimetric;
        if (info.Format == OfficeImageFormat.Tiff) {
            return OfficeTiffCodec.TryDecodePage(encoded, effective.FrameIndex, effective, out image, profile, intent);
        }
        if (info.Format == OfficeImageFormat.Jpeg) {
            // Reserve RGBA plus a possible oriented copy while JPEG budgets coefficients and samples.
            if (!OfficeJpegCodec.TryDecodeColorComponents(encoded, null, false, out var samples,
                    out int width, out int height, out int channels, cancellationToken: effective.CancellationToken,
                    retainedManagedBytes: checked(effective.RetainedManagedBytes + (long)pixels * 8)) ||
                width != info.Width || height != info.Height || channels != profile.ComponentCount ||
                (channels != 1 && channels != 3 && channels != 4)) return false;
            byte[] rgba = OfficeRasterGuards.AllocateRgba32(width, height, "ICC image exceeds the raster limit.");
            if (!ConvertImageSamples(samples, channels, rgba, profile, effective, false)) return false;
            if (OfficeImageOrientationNormalizer.TryRead(encoded, effective.CancellationToken, out var orientation)) {
                rgba = OfficeRasterOrientation.Apply(rgba, ref width, ref height, (int)orientation,
                    effective.CancellationToken, "ICC image orientation exceeds the raster limit.");
            }
            image = OfficeRasterImage.FromOwnedRgba32(width, height, rgba);
            return true;
        }
        if (info.Format != OfficeImageFormat.Png || encoded.Length < 26) return false;
        bool gray = encoded[25] == 0 || encoded[25] == 4;
        bool rgb = encoded[25] == 2 || encoded[25] == 3 || encoded[25] == 6;
        if (!(gray && profile.ComponentCount == 1 || rgb && profile.ComponentCount == 3)) return false;
        if (!OfficeRasterImageDecoder.TryDecode(encoded, effective, out var decoded, out _) || decoded == null) return false;
        // The decoder returns straight RGBA; retain alpha and convert the RGB channels in place.
        if (!ConvertImageSamples(decoded.PixelBuffer, 4, decoded.PixelBuffer, profile, effective, true)) return false;
        image = decoded;
        return true;
    }

    private static bool ConvertImageSamples(byte[] samples, int stride, byte[] rgba,
        OfficeIccColorProfile profile, OfficeRasterDecodeOptions options, bool preserveAlpha) {
        var channels = new double[profile.ComponentCount];
        for (int source = 0, target = 0; source < samples.Length; source += stride, target += 4) {
            if ((target & 16383) == 0) options.CancellationToken.ThrowIfCancellationRequested();
            for (int c = 0; c < channels.Length; c++) channels[c] = samples[source + c] / 255D;
            if (!profile.TryConvert(channels, OfficeIccRenderingIntent.RelativeColorimetric, out var color)) return false;
            rgba[target] = color.R; rgba[target + 1] = color.G; rgba[target + 2] = color.B;
            if (!preserveAlpha) rgba[target + 3] = 255;
        }
        return true;
    }
}
