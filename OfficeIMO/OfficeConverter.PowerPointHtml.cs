using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task<OfficeConversionReport> ConvertPowerPointHtmlAsync(PowerPointDocument document,
        Stream output, OfficeConversionOptions options, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var converted = document.ToHtmlResult(options.PowerPointHtml);
        long slides = options.PowerPointHtml?.IncludeHiddenSlides == true ? document.Slides.Count : document.Slides.Count(slide => !slide.Hidden);
        var report = new OfficeConversionReport(options.SourceName!, "Html", itemsWritten: slides,
            diagnostics: converted.Report.Diagnostics.Select(diagnostic => new OfficeDiagnostic(diagnostic.Code,
                diagnostic.Severity switch {
                    HtmlDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                    HtmlDiagnosticSeverity.Warning => OfficeDiagnosticSeverity.Warning,
                    _ => OfficeDiagnosticSeverity.Information
                }, diagnostic.Message, diagnostic.Provenance.SourceAddress, diagnostic.Component, diagnostic.LossKind)));
        if (!converted.Succeeded) throw new OfficeConversionException("PowerPoint HTML conversion failed.", report);
        if (options.FailOnLoss) report.ThrowIfLoss();
        await OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report,
            () => WriteTextAsync(output, converted.Value, token)).ConfigureAwait(false);
        return report;
    }
}
