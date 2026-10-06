using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private OfficeRasterDecodeOptions ImageDecodeOptions() => new() {
        MaximumDecodedPixels = 4_000_000, CancellationToken = _token, RetainedManagedBytes = _profileAllowance,
        IgnoreTiffOrientation = true
    };

    // ECMA-388 15.3.7: a usable associated profile takes precedence; otherwise
    // a usable embedded profile is required. Do not report discarded profile
    // metadata as paint loss when the mandated fallback retains the image colors.
    private bool TryPrepareImageColor(byte[] bytes, OfficeImageFormat format, string part, string name,
        string? associatedUri, out OfficeRasterImage? raster) {
        raster = null;
        string? unusable = null;
        if (associatedUri != null) {
            var associated = ColorProfile(part, associatedUri, reportUnsupported: false);
            if (associated != null && OfficeIccRasterConverter.TryDecodeToSrgb(bytes, associated, ImageDecodeOptions(), out raster)) return true;
            unusable = associated == null ? "Unsupported ICC profile" : "Unsupported image encoding or ICC channel configuration";
        }

        byte[]? embedded = OfficeImageMetadataInspector.ReadIccProfile(bytes, format, 4 * 1024 * 1024, _token, out bool hasProfile);
        if (hasProfile) {
            if (embedded == null) { Loss("Unusable embedded ICC profile"); return false; }
            var profile = ParseColorProfile("embedded:" + name, embedded, reportUnsupported: false);
            if (profile != null && OfficeIccRasterConverter.TryDecodeToSrgb(bytes, profile, ImageDecodeOptions(), out raster)) return true;
            unusable = profile == null ? "Unsupported ICC profile" : "Unsupported image encoding or ICC channel configuration";
        }
        // Preserve the strict error policy when no usable profile remains. A
        // malformed/unusable profile cannot silently fall through to device RGB.
        if (unusable != null) { Loss(unusable); return false; }
        var metadata = OfficeImageMetadataInspector.Inspect(bytes, format, _profileAllowance, _token);
        // ECMA-388 permits an explicit error when no usable device profile exists.
        if (metadata.HasDeviceMultichannel) { Loss("N-channel image requires a usable ICC profile"); return false; }
        if (metadata.HasDeviceCmyk) { Loss("CMYK image requires a usable ICC profile"); return false; }
        if (metadata.HasColorRenderingMetadata || metadata.HasNonSrgbPngCalibration) {
            // ECMA-388 M8.30 defaults integer PNG samples to sRGB without a
            // usable ICC profile. Decode and re-encode to remove gAMA/cHRM so
            // downstream SVG/PDF consumers cannot reinterpret those samples.
            if (format == OfficeImageFormat.Png && !metadata.HasOtherPngColorRenderingMetadata &&
                OfficeRasterImageDecoder.TryDecodePngDefault(bytes, ImageDecodeOptions(), out raster) && raster != null) return true;
            // JPEG and TIFF use the same integer sRGB/gray defaults (M8.30).
            // Decode the supported device samples and remove non-ICC descriptions
            // before an SVG/PDF consumer can apply its own calibration or EXIF rules.
            if ((format == OfficeImageFormat.Jpeg || format == OfficeImageFormat.Tiff || format == OfficeImageFormat.JpegXr) &&
                OfficeRasterImageDecoder.TryDecode(bytes, ImageDecodeOptions(), out raster, out _) && raster != null) return true;
            Loss("Image color metadata without a supported ICC profile"); return false;
        }
        if (format == OfficeImageFormat.Png && metadata.HasPngAnimation &&
            !OfficeRasterImageDecoder.TryDecodePngDefault(bytes, ImageDecodeOptions(), out raster)) {
            Loss("PNG static image could not be decoded within the native resource limits"); return false;
        }
        // SVG consumers commonly support only eight-bit DCT JPEG. Preserve the
        // managed sample projection when the source uses lossless/high precision.
        bool normalizeJpeg = format == OfficeImageFormat.Jpeg &&
            (metadata.JpegSamplePrecision != 8 || metadata.JpegFrameMarker is not (0xC0 or 0xC1 or 0xC2));
        if ((format == OfficeImageFormat.Tiff || format == OfficeImageFormat.JpegXr || normalizeJpeg) &&
            (!OfficeRasterImageDecoder.TryDecode(bytes, ImageDecodeOptions(), out raster, out _) || raster == null)) {
            Loss("Unsupported image encoding or ICC channel configuration"); return false;
        }
        return true;
    }
}
