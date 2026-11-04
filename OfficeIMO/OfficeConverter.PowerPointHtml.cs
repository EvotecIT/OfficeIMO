using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task<OfficeConversionReport> ConvertPowerPointHtmlAsync(PowerPointDocument document,
        Stream output, OfficeConversionOptions options, CancellationToken token) {
        string html = document.ToHtml(options.PowerPointHtml, out OfficeConversionReport stage, token);
        long slides = options.PowerPointHtml?.IncludeHiddenSlides == true ? document.Slides.Count : document.Slides.Count(slide => !slide.Hidden);
        var report = new OfficeConversionReport(options.SourceName!, "Html", itemsWritten: slides, diagnostics: stage.Diagnostics);
        if (options.FailOnLoss) report.ThrowIfLoss();
        await OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report,
            () => WriteTextAsync(output, html, token)).ConfigureAwait(false);
        return report;
    }
}
