using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    internal long RetainedEncodingBytes => checked(RetainedProfileBytes * 4L + 65536L);

    /// <summary>Projects and applies this metadata to an already encoded managed raster container.</summary>
    /// <remarks>This operation lets animation encoders share the same metadata policy as still encoders without a dependency
    /// on the animation engine. JPEG, PNG, WebP, TIFF, BMP, and GIF support primary-image metadata replacement;
    /// portable maps, TGA, and icons report omitted profile families. The caller's bytes and metadata remain independent
    /// of the returned mutable array and metadata snapshot. Cancellation propagates; no partial result is returned.
    /// Existing 128 MiB encoded and 256 MiB managed working-memory limits apply. The requested ceiling limits returned
    /// bytes; metadata rewrite allocations remain subject to the existing global limits.</remarks>
    /// <param name="encodedBytes">A complete encoded raster container, including externally encoded GIF or APNG animation.</param>
    /// <param name="options">Resolution override and density-writing policy only; compression settings cannot alter already encoded pixels.
    /// The shared write switch suppresses native and Exif density. PNG's switch suppresses pHYs, and JPEG's switch suppresses JFIF;
    /// both retain supplied Exif density. TIFF and WebP store density in Exif, so their switches suppress those density tags.</param>
    /// <param name="maximumEncodedBytes">Positive ceiling for the complete rewritten container, at most 128 MiB.</param>
    /// <param name="cancellationToken">Cancellation observed during inspection and rewriting.</param>
    /// <returns>Owned rewritten bytes, the applied projection, and unsupported supplied profile-family omissions.</returns>
    public OfficeRasterEncodingResult ApplyForEncoding(byte[] encodedBytes, OfficeRasterEncodingOptions? options = null, long maximumEncodedBytes = 134217728L,
        CancellationToken cancellationToken = default) => ApplyForEncoding(encodedBytes, options, maximumEncodedBytes, cancellationToken, 0L);

    internal OfficeRasterEncodingResult ApplyForEncoding(byte[] encodedBytes, OfficeRasterEncodingOptions? options, long maximumEncodedBytes,
        CancellationToken cancellationToken, long additionallyRetainedBytes) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        ValidateEncodingCeiling(maximumEncodedBytes);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterEncodingOptions? effectiveOptions = options?.Clone();
        CheckEncodingCeiling(encodedBytes.LongLength, maximumEncodedBytes);
        if (!OfficeImageReader.TryIdentifyByContent(encodedBytes, null, cancellationToken, out OfficeImageInfo info)) throw new FormatException("The encoded raster container is malformed or unsupported.");
        OfficeImageMetadata requested = Clone();
        if (effectiveOptions?.Resolution != null) requested.Resolution = effectiveOptions.Resolution;
        OfficeImageMetadata projected = requested.PrepareForEncoding(info.Format, out OfficeImageMetadataProfileKinds omitted);
        bool writeDensity = WritesEncodingDensity(info.Format, effectiveOptions);
        bool writeExifDensity = effectiveOptions?.WriteResolutionMetadata != false &&
            (info.Format != OfficeImageFormat.Tiff || effectiveOptions?.Tiff.WriteResolution != false) &&
            (info.Format != OfficeImageFormat.Webp || effectiveOptions?.Webp.WritePhysicalResolution != false);
        if (!writeExifDensity) {
            projected.RemoveExifValue(OfficeExifTag.XResolution);
            projected.RemoveExifValue(OfficeExifTag.YResolution);
            projected.RemoveExifValue(OfficeExifTag.ResolutionUnit);
        }
        byte[] output;
        if (info.Format == OfficeImageFormat.PortableMap || info.Format == OfficeImageFormat.Tga || info.Format == OfficeImageFormat.Icon) {
            if (checked(encodedBytes.LongLength * 2L + additionallyRetainedBytes + RetainedEncodingBytes) > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Metadata encoding exceeds the managed working-memory limit.");
            output = (byte[])encodedBytes.Clone();
        } else {
            output = ApplyCore(encodedBytes, projected, cancellationToken, checked(additionallyRetainedBytes + RetainedEncodingBytes), writeDensity);
        }
        cancellationToken.ThrowIfCancellationRequested();
        CheckEncodingCeiling(output.LongLength, maximumEncodedBytes);
        return new OfficeRasterEncodingResult(output, projected, omitted);
    }

    internal static void ValidateEncodingCeiling(long maximumEncodedBytes) {
        if (maximumEncodedBytes < 1L || maximumEncodedBytes > OfficeRasterGuards.MaximumEncodedBytes) throw new ArgumentOutOfRangeException(nameof(maximumEncodedBytes), "The encoding ceiling must be from one byte through 128 MiB.");
    }

    private static bool WritesEncodingDensity(OfficeImageFormat format, OfficeRasterEncodingOptions? options) {
        if (options == null) return true;
        if (!options.WriteResolutionMetadata) return false;
        return format switch {
            OfficeImageFormat.Png => options.Png.WritePhysicalResolution,
            OfficeImageFormat.Jpeg => options.Jpeg.WriteJfifHeader,
            OfficeImageFormat.Tiff => options.Tiff.WriteResolution,
            OfficeImageFormat.Webp => options.Webp.WritePhysicalResolution,
            _ => true
        };
    }

    internal static void CheckEncodingCeiling(long actual, long maximumEncodedBytes) {
        if (actual > maximumEncodedBytes) throw new OfficeImageExportBatchLimitException(nameof(maximumEncodedBytes), actual, maximumEncodedBytes);
    }
}
