using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

/// <summary>Direct HTML image-export helpers backed by the shared HTML render scene.</summary>
public static partial class HtmlImageExportExtensions {
    internal static OfficeImageExportResult ExportImage(this IHtmlDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        return HtmlRenderEngine.ExecuteWithDeadline(resolved, CancellationToken.None, operationCancellationToken => {
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved, operationCancellationToken);
            if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
            return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.DiagnosticReport, operationCancellationToken);
        });
    }

    internal static IReadOnlyList<OfficeImageExportResult> ExportImages(this IHtmlDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null) {
        HtmlRenderOptions resolved = Normalize(options, 0);
        return HtmlRenderEngine.ExecuteWithDeadline(resolved, CancellationToken.None, operationCancellationToken => {
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved, operationCancellationToken);
            var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
            foreach (HtmlRenderPage page in rendered.Pages) {
                results.Add(RenderPage(page, format, resolved, rendered.DiagnosticReport, operationCancellationToken));
            }
            return results.AsReadOnly();
        });
    }

    internal static async Task<OfficeImageExportResult> ExportImageAsync(this IHtmlDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        return await HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken, async operationCancellationToken => {
            HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, operationCancellationToken).ConfigureAwait(false);
            operationCancellationToken.ThrowIfCancellationRequested();
            if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
            return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.DiagnosticReport, operationCancellationToken);
        }).ConfigureAwait(false);
    }

    internal static async Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this IHtmlDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, 0);
        return await HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken, async operationCancellationToken => {
            HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, operationCancellationToken).ConfigureAwait(false);
            var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
            foreach (HtmlRenderPage page in rendered.Pages) {
                operationCancellationToken.ThrowIfCancellationRequested();
                results.Add(RenderPage(page, format, resolved, rendered.DiagnosticReport, operationCancellationToken));
            }
            return (IReadOnlyList<OfficeImageExportResult>)results.AsReadOnly();
        }).ConfigureAwait(false);
    }

    private static OfficeImageExportResult RenderPage(HtmlRenderPage page, OfficeImageExportFormat format, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeDrawing drawing = page.CreateDrawing(cancellationToken);
        var exportDiagnostics = new List<OfficeImageExportDiagnostic>(MapDiagnostics(diagnostics));
        string source = "HTML render page " + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        drawing.AppendFontDiagnostics(exportDiagnostics, source);
        var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, exportDiagnostics, source);
        byte[] bytes;
        int width;
        int height;
        if (format == OfficeImageExportFormat.Svg) {
            bytes = OfficeDrawingSvgExporter.ToSvgBytes(
                drawing,
                options.Scale,
                OfficeSvgSizeUnit.Pixel,
                fallbackCodec,
                resourceIdPrefix: null,
                cancellationToken: cancellationToken);
            width = Math.Max(1, (int)Math.Ceiling(page.Width * options.Scale));
            height = Math.Max(1, (int)Math.Ceiling(page.Height * options.Scale));
        } else if (format.IsRaster()) {
            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                drawing.Width,
                drawing.Height,
                format,
                options,
                source);
            if (plan.Diagnostic != null) exportDiagnostics.Add(plan.Diagnostic);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                Scale = plan.Limit.Scale,
                Background = options.BackgroundColor,
                ImageCodec = fallbackCodec,
                TextShapingProvider = options.TextShapingProvider,
                TextShapingLanguage = options.TextShapingLanguage,
                DiagnosticSink = exportDiagnostics,
                DiagnosticSource = source,
                CancellationToken = cancellationToken
            });
            bytes = OfficeRasterImageEncoder.Encode(
                image,
                format,
                plan.CreateEncodingOptions());
            width = image.Width;
            height = image.Height;
        } else {
            throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image export format.");
        }
        cancellationToken.ThrowIfCancellationRequested();
        return options.EnsureAccepted(new OfficeImageExportResult(
            format,
            width,
            height,
            bytes,
            "Page " + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            source,
            exportDiagnostics));
    }

    private static IReadOnlyList<OfficeImageExportDiagnostic> MapDiagnostics(HtmlDiagnosticReport report) {
        var diagnostics = new List<OfficeImageExportDiagnostic>(report.Count);
        foreach (HtmlDiagnostic diagnostic in report.Diagnostics) {
            OfficeImageExportDiagnosticSeverity severity = diagnostic.Severity == HtmlDiagnosticSeverity.Error
                ? OfficeImageExportDiagnosticSeverity.Error
                : diagnostic.Severity == HtmlDiagnosticSeverity.Warning ? OfficeImageExportDiagnosticSeverity.Warning : OfficeImageExportDiagnosticSeverity.Info;
            diagnostics.Add(new OfficeImageExportDiagnostic(
                severity,
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Source,
                diagnostic.LossKind switch {
                    OfficeConversionLossKind.Approximation => OfficeConversionLossKind.Approximation,
                    OfficeConversionLossKind.Omission => OfficeConversionLossKind.Omission,
                    OfficeConversionLossKind.Failure => OfficeConversionLossKind.Failure,
                    _ => OfficeConversionLossKind.None
                }));
        }
        return diagnostics.AsReadOnly();
    }

    private static HtmlRenderOptions Normalize(HtmlRenderOptions? options, int pageIndex) {
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.Validate();
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex), "HTML render page index cannot be negative.");
        return resolved;
    }

}
