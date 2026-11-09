using OfficeIMO.Drawing;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    /// <summary>Renders report pages using the encoding and selection in the shared options.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this ProjectView view,
        ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        ToImages(view, options?.Format ?? OfficeImageExportFormat.Png, options, cancellationToken);

    /// <summary>Exports report pages with encoded density, format metadata, diagnostics and bounded aggregate output.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this ProjectView view, OfficeImageExportFormat format,
        ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        ToImages(view, format, results.Add, options, cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams image results under one shared deadline and batch budget. Already delivered pages remain delivered if a later page fails.</summary>
    public static void ToImages(this ProjectView view, OfficeImageExportFormat format, OfficeImageExportConsumer consumer,
        ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view); ArgumentNullException.ThrowIfNull(consumer);
        var effective = options?.Snapshot() ?? new ProjectImageExportOptions();
        OfficeImageExportBatchProcessor.Run(effective, (emit, token) => {
            var typography = new OfficeRenderingProfile("project-images", effective.Fonts, effective.TextShapingProvider,
                effective.TextShapingLanguage, effective.ImageCodec, effective.Policy);
            var pages = view.Render(typography, token);
            IReadOnlyList<int> selected = effective.PageSelector?.Resolve(pages.Count, effective.MaximumOutputCount)
                ?? Enumerable.Range(1, pages.Count).ToArray();
            if (selected.Count > effective.MaximumOutputCount) throw new OfficeImageExportBatchLimitException(nameof(effective.MaximumOutputCount), selected.Count, effective.MaximumOutputCount);
            for (int index = 0; index < selected.Count; index++) {
                token.ThrowIfCancellationRequested();
                int number = selected[index];
                var image = pages[number - 1].Drawing.ToImage(format, effective, token).WithIdentity("Page " + number, "Project").WithSequence(index, selected.Count);
                emit(effective.EnsureAccepted(OfficeImageExportEvidence.AttachSourceReport(image, view.Report)));
            }
        }, consumer, cancellationToken);
    }
}
