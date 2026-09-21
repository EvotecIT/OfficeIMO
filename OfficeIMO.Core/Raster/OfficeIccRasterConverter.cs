using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Reason a packed ICC raster conversion did not produce sRGB pixels.</summary>
public enum OfficeIccRasterConversionStatus {
    /// <summary>Every source pixel was converted.</summary>
    Converted,
    /// <summary>The profile exceeds the requested parser input limit.</summary>
    ProfileLimitExceeded,
    /// <summary>The image dimensions exceed the requested pixel limit.</summary>
    PixelLimitExceeded,
    /// <summary>The source, parsed profile, and destination cannot fit the declared managed budget.</summary>
    AllocationLimitExceeded,
    /// <summary>The profile is malformed or has no supported device-to-PCS transform.</summary>
    UnsupportedProfile,
    /// <summary>The source buffer is not tightly packed 8-bit samples for the profile's components.</summary>
    InvalidSamples,
    /// <summary>A supported profile could not convert at least one source pixel.</summary>
    ConversionFailed
}

/// <summary>Limits and rendering intent for explicit packed-sample ICC conversion.</summary>
public sealed class OfficeIccRasterConversionOptions {
    private int _maximumProfileBytes = 4 * 1024 * 1024;
    private long _maximumPixels = 4_000_000L;
    private long _maximumManagedBytes = OfficeRasterGuards.MaximumDecodedBytes;

    /// <summary>Maximum accepted profile length; the hard ceiling is 4 MiB.</summary>
    public int MaximumProfileBytes {
        get => _maximumProfileBytes;
        set {
            if (value < 1 || value > 4 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaximumProfileBytes));
            _maximumProfileBytes = value;
        }
    }

    /// <summary>Maximum number of source pixels; the hard ceiling is 50 million.</summary>
    public long MaximumPixels {
        get => _maximumPixels;
        set {
            if (value < 1L || value > OfficeRasterGuards.MaximumPixels) throw new ArgumentOutOfRangeException(nameof(MaximumPixels));
            _maximumPixels = value;
        }
    }

    /// <summary>Maximum accounted source, profile, parser, and destination bytes; the hard ceiling is 256 MiB.</summary>
    public long MaximumManagedBytes {
        get => _maximumManagedBytes;
        set {
            if (value < 1L || value > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentOutOfRangeException(nameof(MaximumManagedBytes));
            _maximumManagedBytes = value;
        }
    }

    /// <summary>ICC rendering intent applied to every device sample.</summary>
    public OfficeIccRenderingIntent RenderingIntent { get; set; } = OfficeIccRenderingIntent.RelativeColorimetric;

    /// <summary>Cancellation checked during the pixel loop.</summary>
    public CancellationToken CancellationToken { get; set; }
}

/// <summary>Converts bounded, tightly packed 8-bit RGB or CMYK device samples to opaque sRGB.</summary>
/// <remarks>
/// This API accepts raw device samples and ICC bytes supplied by the caller. It does not extract
/// metadata from encoded images, alter the source buffer, or change image optimizer metadata reports.
/// Alpha, planar samples, higher bit depths, and implicit device-color conversion are outside this contract.
/// </remarks>
public static class OfficeIccRasterConverter {
    /// <summary>Attempts one complete conversion, returning a typed rejection instead of partial pixels.</summary>
    public static OfficeIccRasterConversionStatus TryConvertToSrgb(
        byte[] deviceSamples,
        int width,
        int height,
        byte[] profileBytes,
        OfficeIccRasterConversionOptions? options,
        out OfficeRasterImage? image) {
        if (deviceSamples == null) throw new ArgumentNullException(nameof(deviceSamples));
        if (profileBytes == null) throw new ArgumentNullException(nameof(profileBytes));
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        OfficeIccRasterConversionOptions effective = options ?? new OfficeIccRasterConversionOptions();
        if (effective.RenderingIntent < OfficeIccRenderingIntent.Perceptual ||
            effective.RenderingIntent > OfficeIccRenderingIntent.AbsoluteColorimetric) {
            throw new ArgumentOutOfRangeException(nameof(options), "Rendering intent is unsupported.");
        }
        image = null;
        effective.CancellationToken.ThrowIfCancellationRequested();
        if (profileBytes.Length == 0 || profileBytes.Length > effective.MaximumProfileBytes) {
            return OfficeIccRasterConversionStatus.ProfileLimitExceeded;
        }
        long pixels = (long)width * height;
        if (pixels > effective.MaximumPixels || pixels > OfficeRasterGuards.MaximumPixels) {
            return OfficeIccRasterConversionStatus.PixelLimitExceeded;
        }
        // Profile parsing can expand LUT entries into doubles. Reserve a conservative upper bound
        // before invoking the parser so a caller cannot bypass the managed allocation policy.
        long outputBytes = checked(pixels * 4L);
        long parserAllowance = checked((long)profileBytes.Length * 32L + 4096L);
        if (ExceedsBudget(deviceSamples.LongLength, profileBytes.LongLength, outputBytes,
                parserAllowance, effective.MaximumManagedBytes)) {
            return OfficeIccRasterConversionStatus.AllocationLimitExceeded;
        }
        if (!OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile) || profile == null) {
            return OfficeIccRasterConversionStatus.UnsupportedProfile;
        }
        int channels = profile.ComponentCount;
        if (channels != 3 && channels != 4) {
            return OfficeIccRasterConversionStatus.UnsupportedProfile;
        }
        if (deviceSamples.LongLength != pixels * channels) {
            return OfficeIccRasterConversionStatus.InvalidSamples;
        }
        if (ExceedsBudget(deviceSamples.LongLength, profileBytes.LongLength, outputBytes,
                Math.Max(parserAllowance, profile.RetainedByteCount), effective.MaximumManagedBytes)) {
            return OfficeIccRasterConversionStatus.AllocationLimitExceeded;
        }
        var converted = new byte[checked((int)outputBytes)];
        var components = new double[channels];
        for (long pixel = 0L; pixel < pixels; pixel++) {
            if ((pixel & 4095L) == 0L) effective.CancellationToken.ThrowIfCancellationRequested();
            int source = checked((int)(pixel * channels));
            for (int channel = 0; channel < channels; channel++) {
                components[channel] = deviceSamples[source + channel] / 255D;
            }
            if (!profile.TryConvert(components, effective.RenderingIntent, out OfficeColor color)) {
                return OfficeIccRasterConversionStatus.ConversionFailed;
            }
            int target = checked((int)(pixel * 4L));
            converted[target] = color.R;
            converted[target + 1] = color.G;
            converted[target + 2] = color.B;
            converted[target + 3] = 255;
        }
        image = OfficeRasterImage.FromOwnedRgba32(width, height, converted);
        return OfficeIccRasterConversionStatus.Converted;
    }

    private static bool ExceedsBudget(long source, long profile, long output, long parser, long maximum) {
        try {
            return checked(source + profile + output + parser) > maximum;
        } catch (OverflowException) {
            return true;
        }
    }
}
