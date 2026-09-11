using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

/// <summary>Optional report adapters that compose Project snapshots with the existing document and rendering owners.</summary>
public static partial class ProjectReportWorkflow {
    /// <summary>Exports one SVG per report page using the Core vector renderer.</summary>
    public static IReadOnlyList<string> ToSvg(ProjectView view, OfficeRenderingProfile? typography = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        return RenderPages(view, typography, cancellationToken).Select((page, index) => {
            cancellationToken.ThrowIfCancellationRequested();
            return OfficeDrawingSvgExporter.ToSvg(page.Drawing, 1, OfficeSvgSizeUnit.Pixel, typography?.ImageCodec, "project-page-" + index + "-", cancellationToken);
        }).ToArray();
    }

    /// <summary>Exports one PNG per page. Supply a typography profile for portable font coverage; raster settings control scale, shaping, diagnostics and pixel limits.</summary>
    public static IReadOnlyList<byte[]> ToPng(ProjectView view, OfficeDrawingRasterRenderOptions? options = null, OfficeRenderingProfile? typography = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        options ??= new OfficeDrawingRasterRenderOptions();
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, options.CancellationToken);
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var effective = new OfficeDrawingRasterRenderOptions { Scale = options.Scale, Background = options.Background, ImageCodec = options.ImageCodec ?? typography?.ImageCodec,
            ThrowOnImageDecodeFailure = options.ThrowOnImageDecodeFailure, TextShapingProvider = options.TextShapingProvider ?? typography?.TextShapingProvider,
            TextShapingLanguage = options.TextShapingLanguage ?? typography?.TextShapingLanguage, DiagnosticSink = diagnostics, DiagnosticSource = options.DiagnosticSource,
            MaximumRasterPixels = options.MaximumRasterPixels, CancellationToken = linked.Token };
        return RenderPages(view, typography, linked.Token).Select(page => {
            cancellationToken.ThrowIfCancellationRequested();
            diagnostics.Clear();
            byte[] bytes = OfficeDrawingRasterRenderer.ToPng(page.Drawing, effective);
            if (options.DiagnosticSink != null) foreach (var diagnostic in diagnostics) options.DiagnosticSink.Add(diagnostic);
            typography?.Policy.EnsureAccepted(diagnostics);
            return bytes;
        }).ToArray();
    }

    /// <summary>Creates a paginated vector PDF through OfficeIMO.Pdf. Typography and compliance follow the supplied PDF options.</summary>
    public static byte[] ToPdf(ProjectView view, PdfOptions? options = null, OfficeRenderingProfile? typography = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        var pages = RenderPages(view, typography, cancellationToken);
        var document = PdfDocument.Create(options);
        document.Compose(builder => {
            if (typography != null) builder.Typography(typography);
            foreach (var page in pages) {
                cancellationToken.ThrowIfCancellationRequested();
                var drawing = page.Drawing;
                builder.Page(p => p.Size(drawing.Width, drawing.Height).Margin(0)
                    .Canvas(canvas => canvas.Drawing(drawing, 0, 0, drawing.Width, drawing.Height)));
            }
        });
        return document.ToBytes(cancellationToken);
    }

    private static IReadOnlyList<ProjectViewPage> RenderPages(ProjectView view, OfficeRenderingProfile? typography, CancellationToken token) {
        var pages = view.Render(token);
        if (typography != null) foreach (var page in pages) page.Drawing.Fonts.AddRange(typography.Fonts);
        return pages;
    }
}
