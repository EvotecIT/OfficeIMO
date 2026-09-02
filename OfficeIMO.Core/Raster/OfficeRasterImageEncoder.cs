using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free encoder for raster export formats.
/// </summary>
public static partial class OfficeRasterImageEncoder {
    internal const double PngMinimumDpi = 0.0127D;
    internal const double JpegMinimumDpi = 0.5D;
    internal const double TiffMinimumDpi = 0.001D;
    internal const double WebpMinimumDpi = 0.0001D;
    internal const double PngMaximumDpi = uint.MaxValue * 0.0254D;
    internal const double JpegMaximumDpi = ushort.MaxValue;
    internal const double TiffMaximumDpi = 1000000D;
    internal const double WebpMaximumDpi = 1000000D;
    internal const int JpegMaximumDimension = ushort.MaxValue;
    internal const int WebpMaximumDimension = 16384;

    /// <summary>Returns the maximum supported pixel width or height for a raster format.</summary>
    public static int GetMaximumDimension(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => int.MaxValue,
        OfficeImageExportFormat.Jpeg => JpegMaximumDimension,
        OfficeImageExportFormat.Tiff => int.MaxValue,
        OfficeImageExportFormat.Webp => WebpMaximumDimension,
        OfficeImageExportFormat.Svg => throw new ArgumentException("SVG output does not have a raster dimension limit.", nameof(format)),
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    /// <summary>Returns the maximum source pixel count accepted by a raster encoder.</summary>
    public static long GetMaximumPixelCount(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => long.MaxValue,
        OfficeImageExportFormat.Jpeg => OfficeRasterGuards.MaximumEncodedBytes / 4L,
        OfficeImageExportFormat.Tiff => long.MaxValue,
        OfficeImageExportFormat.Webp => long.MaxValue,
        OfficeImageExportFormat.Svg => throw new ArgumentException("SVG output does not have a raster pixel limit.", nameof(format)),
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    internal static double GetMinimumDpi(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => PngMinimumDpi,
        OfficeImageExportFormat.Jpeg => JpegMinimumDpi,
        OfficeImageExportFormat.Tiff => TiffMinimumDpi,
        OfficeImageExportFormat.Webp => WebpMinimumDpi,
        OfficeImageExportFormat.Svg => throw new ArgumentException("SVG output does not encode raster density.", nameof(format)),
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    internal static double GetMaximumDpi(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => PngMaximumDpi,
        OfficeImageExportFormat.Jpeg => JpegMaximumDpi,
        OfficeImageExportFormat.Tiff => TiffMaximumDpi,
        OfficeImageExportFormat.Webp => WebpMaximumDpi,
        OfficeImageExportFormat.Svg => throw new ArgumentException("SVG output does not encode raster density.", nameof(format)),
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    internal static double NormalizeDpi(OfficeImageExportFormat format, double dpi) {
        if (double.IsNaN(dpi) || double.IsInfinity(dpi) || dpi <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(dpi), "Raster DPI must be finite and greater than zero.");
        }
        double normalized = Math.Min(GetMaximumDpi(format), Math.Max(GetMinimumDpi(format), dpi));
        return format == OfficeImageExportFormat.Jpeg
            ? Math.Round(normalized, MidpointRounding.AwayFromZero)
            : normalized;
    }

    /// <summary>Encodes an RGBA image using the requested raster format.</summary>
    public static byte[] Encode(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        OfficeRasterEncodingOptions? options = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterEncodingOptions effective =
            (options ?? new OfficeRasterEncodingOptions()).Resolve(format);
        return format switch {
            OfficeImageExportFormat.Png => OfficePngWriter.Encode(
                image,
                effective.Png ?? throw new InvalidOperationException("PNG encoding options cannot be null.")),
            OfficeImageExportFormat.Jpeg => OfficeJpegCodec.Encode(
                image,
                effective.Jpeg ?? throw new InvalidOperationException("JPEG encoding options cannot be null.")),
            OfficeImageExportFormat.Tiff => OfficeTiffCodec.Encode(
                image,
                effective.Tiff ?? throw new InvalidOperationException("TIFF encoding options cannot be null.")),
            OfficeImageExportFormat.Webp => effective.WriteResolutionMetadata
                ? OfficeWebpCodec.Encode(image, effective.DpiX, effective.DpiY)
                : OfficeWebpCodec.Encode(image),
            OfficeImageExportFormat.Svg => throw new ArgumentException("SVG output requires a vector renderer.", nameof(format)),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };
    }

    /// <summary>
    /// Encodes an RGBA image while observing cancellation and enforcing an encoded-byte ceiling
    /// as bytes are produced.
    /// </summary>
    public static byte[] Encode(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        OfficeRasterEncodingOptions? options,
        long maximumEncodedBytes,
        CancellationToken cancellationToken = default) {
        var budget = new OfficeImageExportEncodingBudget(maximumEncodedBytes);
        return Encode(image, format, options, budget, cancellationToken);
    }

    internal static byte[] Encode(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        OfficeRasterEncodingOptions? options,
        OfficeImageExportEncodingBudget budget,
        CancellationToken cancellationToken) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        if (budget == null) throw new ArgumentNullException(nameof(budget));
        OfficeRasterEncodingOptions effective =
            (options ?? new OfficeRasterEncodingOptions()).Resolve(format);
        using var output = new OfficeImageExportEncodingMemoryStream(
            budget,
            cancellationToken,
            GetRetainedManagedBytes(image, format, effective));
        EncodeToCore(image, format, output, effective, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return output.ToBoundedArray();
    }

    private static long GetRetainedManagedBytes(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        OfficeRasterEncodingOptions options) {
        long retained = checked(image.PixelBuffer.LongLength + 24L);
        if (format != OfficeImageExportFormat.Jpeg) return retained;

        OfficeJpegEncodeOptions jpeg = options.Jpeg;
        retained = checked(retained + jpeg.RetainedManagedBytes);
        if (jpeg.Metadata.ExifBuffer != null) retained = checked(retained + jpeg.Metadata.ExifBuffer.LongLength + 24L);
        if (jpeg.Metadata.XmpBuffer != null) retained = checked(retained + jpeg.Metadata.XmpBuffer.LongLength + 24L);
        if (jpeg.Metadata.IccBuffer != null) retained = checked(retained + jpeg.Metadata.IccBuffer.LongLength + 24L);
        return retained;
    }
}
