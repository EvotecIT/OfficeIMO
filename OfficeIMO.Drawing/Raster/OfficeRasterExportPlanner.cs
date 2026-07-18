using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>Resolved raster allocation plan shared by document image exporters.</summary>
public readonly struct OfficeRasterExportPlan {
    internal OfficeRasterExportPlan(
        OfficeRasterScaleLimit limit,
        long maximumPixels,
        int maximumDimension,
        OfficeImageExportDiagnostic? diagnostic) {
        Limit = limit;
        MaximumPixels = maximumPixels;
        MaximumDimension = maximumDimension;
        Diagnostic = diagnostic;
    }

    /// <summary>Effective raster scale and dimensions.</summary>
    public OfficeRasterScaleLimit Limit { get; }

    /// <summary>Effective pixel-count ceiling after combining caller and encoder limits.</summary>
    public long MaximumPixels { get; }

    /// <summary>Effective encoder dimension ceiling.</summary>
    public int MaximumDimension { get; }

    /// <summary>Scale-reduction diagnostic, or null when the request fit unchanged.</summary>
    public OfficeImageExportDiagnostic? Diagnostic { get; }
}

/// <summary>Creates overflow-safe raster allocation plans before any image surface is allocated.</summary>
public static class OfficeRasterExportPlanner {
    /// <summary>
    /// Resolves a raster plan from shared export options and the selected encoder's limits.
    /// </summary>
    public static OfficeRasterExportPlan Resolve(
        double width,
        double height,
        OfficeImageExportFormat format,
        OfficeImageExportOptions options,
        string? source = null) =>
        Resolve(width, height, format, options, options?.MaximumRasterPixels ?? 0L, source);

    /// <summary>
    /// Resolves a raster plan with an additional renderer-specific pixel ceiling.
    /// </summary>
    public static OfficeRasterExportPlan Resolve(
        double width,
        double height,
        OfficeImageExportFormat format,
        OfficeImageExportOptions options,
        long rendererMaximumPixels,
        string? source = null) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (!format.IsRaster()) throw new ArgumentException("A raster output format is required.", nameof(format));
        options.ValidateImageExportOptions();
        if (rendererMaximumPixels < 1L) throw new ArgumentOutOfRangeException(nameof(rendererMaximumPixels));

        long maximumPixels = Math.Min(
            Math.Min(options.MaximumRasterPixels, rendererMaximumPixels),
            OfficeRasterImageEncoder.GetMaximumPixelCount(format));
        int maximumDimension = OfficeRasterImageEncoder.GetMaximumDimension(format);
        OfficeRasterScaleLimit limit = OfficeRasterScaleLimiter.Resolve(
            width,
            height,
            options.Scale,
            maximumPixels,
            maximumDimension);

        if (!limit.WasLimited) {
            return new OfficeRasterExportPlan(limit, maximumPixels, maximumDimension, diagnostic: null);
        }

        if (options.RasterOverflowBehavior == OfficeRasterOverflowBehavior.Throw) {
            throw new OfficeImageExportLimitException(
                options.Scale,
                CalculateRequestedPixels(width, height, options.Scale),
                maximumPixels,
                maximumDimension);
        }

        var diagnostic = new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            OfficeImageExportDiagnosticCodes.RasterScaleReduced,
            "The raster scale was reduced from " + Format(options.Scale) + " to " + Format(limit.Scale) +
            " to satisfy the effective limit of " + maximumPixels.ToString(CultureInfo.InvariantCulture) +
            " pixels and " + maximumDimension.ToString(CultureInfo.InvariantCulture) + " pixels per dimension.",
            source);
        return new OfficeRasterExportPlan(limit, maximumPixels, maximumDimension, diagnostic);
    }

    private static long CalculateRequestedPixels(double width, double height, double scale) {
        double scaledWidth = Math.Ceiling(width * scale);
        double scaledHeight = Math.Ceiling(height * scale);
        if (double.IsNaN(scaledWidth) || double.IsInfinity(scaledWidth) || scaledWidth > int.MaxValue ||
            double.IsNaN(scaledHeight) || double.IsInfinity(scaledHeight) || scaledHeight > int.MaxValue ||
            scaledWidth <= 0D || scaledHeight <= 0D ||
            scaledWidth > long.MaxValue / scaledHeight) {
            return long.MaxValue;
        }
        return (long)scaledWidth * (long)scaledHeight;
    }

    private static string Format(double value) =>
        value.ToString("0.########", CultureInfo.InvariantCulture);
}
