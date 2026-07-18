using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static class VisioImageExportEngine {
    private const double DefaultPixelsPerInch = 96D;
    private const long MaximumSupersampledPixels = 32_000_000L;

    internal static OfficeImageExportResult Render(
        VisioPage page,
        OfficeImageExportFormat format,
        VisioImageExportOptions options,
        string? name = null,
        string? source = null) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();

        double pixelsPerInch = ResolvePixelsPerInch(options.Scale);
        string resultName = string.IsNullOrWhiteSpace(name) ? page.Name : name!;
        string resultSource = string.IsNullOrWhiteSpace(source) ? "Visio page" : source!;
        if (format == OfficeImageExportFormat.Svg) {
            int width = Scaled(page.Width, pixelsPerInch);
            int height = Scaled(page.Height, pixelsPerInch);
            byte[] bytes = Encoding.UTF8.GetBytes(VisioSvgRenderer.Render(page, CreateSvgOptions(options, pixelsPerInch)));
            return new OfficeImageExportResult(
                format,
                width,
                height,
                bytes,
                resultName,
                resultSource);
        }

        if (!format.IsRaster()) {
            throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image export format.");
        }

        var diagnostics = new List<OfficeImageExportDiagnostic>();
        long workingPixelLimit = Math.Max(1L, MaximumSupersampledPixels / ((long)options.Supersampling * options.Supersampling));
        long maximumPixels = Math.Min(
            Math.Min(options.MaximumRasterPixels, workingPixelLimit),
            OfficeRasterImageEncoder.GetMaximumPixelCount(format));
        OfficeRasterScaleLimit limit = OfficeRasterScaleLimiter.Resolve(
            Math.Max(page.Width, 0.01D),
            Math.Max(page.Height, 0.01D),
            pixelsPerInch,
            maximumPixels,
            OfficeRasterImageEncoder.GetMaximumDimension(format));
        if (limit.WasLimited) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                "VISIO_IMAGE_RASTER_SCALE_LIMITED",
                "The raster resolution was reduced from " + Format(pixelsPerInch) + " to " + Format(limit.Scale) +
                " pixels per inch to stay within the configured raster limits.",
                resultSource));
        }

        OfficeRasterImage image = VisioPngRenderer.RenderRaster(page, CreatePngOptions(options, limit.Scale));
        byte[] encoded = OfficeRasterImageEncoder.Encode(image, format, options.RasterEncoding);
        return new OfficeImageExportResult(
            format,
            image.Width,
            image.Height,
            encoded,
            resultName,
            resultSource,
            diagnostics);
    }

    private static VisioPngSaveOptions CreatePngOptions(VisioImageExportOptions options, double pixelsPerInch) =>
        new VisioPngSaveOptions {
            PixelsPerInch = pixelsPerInch,
            BackgroundColor = options.BackgroundColor,
            RenderText = options.RenderText,
            FontFilePath = options.FontFilePath,
            FontFaceName = options.FontFaceName,
            FontCollectionIndex = options.FontCollectionIndex,
            RenderStencilArtwork = options.RenderStencilArtwork,
            RenderConnectorLabels = options.RenderConnectorLabels,
            ResolveConnectorLabelOverlaps = options.ResolveConnectorLabelOverlaps,
            Supersampling = options.Supersampling
        };

    private static VisioSvgSaveOptions CreateSvgOptions(VisioImageExportOptions options, double pixelsPerInch) =>
        new VisioSvgSaveOptions {
            PixelsPerInch = pixelsPerInch,
            BackgroundColor = options.BackgroundColor,
            RenderText = options.RenderText,
            RenderStencilArtwork = options.RenderStencilArtwork,
            RenderConnectorLabels = options.RenderConnectorLabels,
            ResolveConnectorLabelOverlaps = options.ResolveConnectorLabelOverlaps,
            IncludeXmlDeclaration = options.IncludeSvgXmlDeclaration
        };

    private static int Scaled(double inches, double pixelsPerInch) {
        double value = Math.Ceiling(Math.Max(inches, 0.01D) * pixelsPerInch);
        if (double.IsNaN(value) || double.IsInfinity(value) || value > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(pixelsPerInch), "The requested SVG dimensions exceed supported integer bounds.");
        }
        return Math.Max(1, (int)value);
    }

    private static double ResolvePixelsPerInch(double scale) {
        double pixelsPerInch = DefaultPixelsPerInch * scale;
        if (double.IsNaN(pixelsPerInch) || double.IsInfinity(pixelsPerInch) || pixelsPerInch <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(scale), "The requested scale exceeds supported Visio image dimensions.");
        }
        return pixelsPerInch;
    }

    private static string Format(double value) => value.ToString("0.########", CultureInfo.InvariantCulture);
}
