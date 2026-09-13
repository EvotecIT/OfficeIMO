using OfficeIMO.Drawing;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    /// <summary>Starts a fluent export using the shared OfficeIMO image contracts.</summary>
    public static ProjectImageExportBuilder Images(ProjectView view, ProjectImageExportOptions? options = null) {
        ArgumentNullException.ThrowIfNull(view);
        return new ProjectImageExportBuilder(view, options?.Snapshot() ?? new ProjectImageExportOptions());
    }

    /// <summary>Exports report pages with encoded density, format metadata, diagnostics and bounded aggregate output.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(ProjectView view, OfficeImageExportFormat format = OfficeImageExportFormat.Png,
        ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        ExportImages(view, format, results.Add, options, cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams image results under one shared deadline and batch budget. Already delivered pages remain delivered if a later page fails.</summary>
    public static void ExportImages(ProjectView view, OfficeImageExportFormat format, OfficeImageExportConsumer consumer,
        ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view); ArgumentNullException.ThrowIfNull(consumer);
        var effective = options?.Snapshot() ?? new ProjectImageExportOptions();
        OfficeImageExportBatchProcessor.Run(effective, (emit, token) => {
            var typography = new OfficeRenderingProfile("project-images", effective.Fonts, effective.TextShapingProvider,
                effective.TextShapingLanguage, effective.ImageCodec, effective.Policy);
            var pages = view.Render(typography, token);
            if (pages.Count > effective.MaximumOutputCount) throw new InvalidOperationException("Report exceeds MaximumOutputCount.");
            foreach (var page in pages) {
                token.ThrowIfCancellationRequested();
                emit(page.Drawing.ExportImage(format, effective, token));
            }
        }, consumer, cancellationToken);
    }
}
