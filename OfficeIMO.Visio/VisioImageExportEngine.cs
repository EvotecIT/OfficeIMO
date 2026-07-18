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
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        if (format == OfficeImageExportFormat.Svg) {
            int width = Scaled(page.Width, pixelsPerInch);
            int height = Scaled(page.Height, pixelsPerInch);
            byte[] bytes = Encoding.UTF8.GetBytes(VisioSvgRenderer.Render(page, CreateSvgOptions(options, pixelsPerInch, diagnostics, resultSource)));
            return new OfficeImageExportResult(
                format,
                width,
                height,
                bytes,
                resultName,
                resultSource,
                diagnostics);
        }

        if (!format.IsRaster()) {
            throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image export format.");
        }

        long workingPixelLimit = Math.Max(1L, MaximumSupersampledPixels / ((long)options.Supersampling * options.Supersampling));
        OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
            Math.Max(page.Width * DefaultPixelsPerInch, 0.01D),
            Math.Max(page.Height * DefaultPixelsPerInch, 0.01D),
            format,
            options,
            workingPixelLimit,
            resultSource);
        if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);

        OfficeRasterImage image = VisioPngRenderer.RenderRaster(
            page,
            CreatePngOptions(options, ResolvePixelsPerInch(plan.Limit.Scale), diagnostics, resultSource));
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

    private static VisioPngSaveOptions CreatePngOptions(
        VisioImageExportOptions options,
        double pixelsPerInch,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        string source) =>
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
            Supersampling = options.Supersampling,
            ImageCodec = options.ImageCodec,
            ImageDiagnostics = diagnostics,
            ImageDiagnosticSource = source
        };

    private static VisioSvgSaveOptions CreateSvgOptions(
        VisioImageExportOptions options,
        double pixelsPerInch,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        string source) =>
        new VisioSvgSaveOptions {
            PixelsPerInch = pixelsPerInch,
            BackgroundColor = options.BackgroundColor,
            RenderText = options.RenderText,
            RenderStencilArtwork = options.RenderStencilArtwork,
            RenderConnectorLabels = options.RenderConnectorLabels,
            ResolveConnectorLabelOverlaps = options.ResolveConnectorLabelOverlaps,
            IncludeXmlDeclaration = options.IncludeSvgXmlDeclaration,
            ImageCodec = options.ImageCodec,
            ImageDiagnostics = diagnostics,
            ImageDiagnosticSource = source
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
}
