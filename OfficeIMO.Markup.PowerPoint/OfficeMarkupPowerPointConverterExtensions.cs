using OfficeIMO.PowerPoint;
using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markup.PowerPoint;

/// <summary>Converts typed OfficeIMO Markup into editable PowerPoint documents.</summary>
public static class OfficeMarkupPowerPointConverterExtensions {
    /// <summary>Converts markup into a detached document owned by the caller.</summary>
    public static PowerPointDocument ToPowerPoint(this OfficeMarkupDocument document, MarkupPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToPowerPoint(options, out _, cancellationToken);

    /// <summary>Converts markup and retains mapping and native validation evidence.</summary>
    public static PowerPointDocument ToPowerPoint(this OfficeMarkupDocument document, MarkupPowerPointExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        cancellationToken.ThrowIfCancellationRequested();
        MarkupPowerPointExportOptions resolved = options ?? new MarkupPowerPointExportOptions();
        IReadOnlyList<OfficeMarkupDiagnostic> diagnostics = CollectDiagnostics(document, resolved);
        var result = new OfficeMarkupPowerPointExporter().Build(document, resolved, diagnostics);
        report = result.Report;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (!result.Succeeded) throw new OfficeConversionException("Markup conversion failed validation.", report);
            if (resolved.FailOnLoss && report.HasLoss) throw new OfficeConversionException("Markup conversion reported fidelity loss.", report);
            return result.Value;
        } catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Converts markup asynchronously and returns the document with its completed report.</summary>
    public static Task<OfficeConversionResult<PowerPointDocument>> ToPowerPointAsync(this OfficeMarkupDocument document,
        MarkupPowerPointExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes converted PowerPoint output and returns its report.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(
        this OfficeMarkupDocument document, string path, MarkupPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PowerPointDocument value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(path));
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Writes converted PowerPoint output and returns its report.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(
        this OfficeMarkupDocument document, Stream stream, MarkupPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PowerPointDocument value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(stream));
        return report;
    }

    /// <summary>Asynchronously writes converted PowerPoint output and returns its report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(
        this OfficeMarkupDocument document, string path, MarkupPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PowerPointDocument value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Asynchronously writes converted PowerPoint output and returns its report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(
        this OfficeMarkupDocument document, Stream stream, MarkupPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PowerPointDocument value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return report;
    }

    private static IReadOnlyList<OfficeMarkupDiagnostic> CollectDiagnostics(
        OfficeMarkupDocument document,
        MarkupPowerPointExportOptions options) {
        var diagnostics = new List<OfficeMarkupDiagnostic>(OfficeMarkupValidator.Validate(document));
        foreach (OfficeMarkupBlock block in document.DescendantsAndSelf()) {
            if (block is OfficeMarkupCodeBlock || block is OfficeMarkupRawMarkdownBlock) {
                diagnostics.Add(Omitted(block));
            } else if (block is OfficeMarkupExtensionBlock) {
                diagnostics.Add(new OfficeMarkupDiagnostic(
                    OfficeMarkupDiagnosticSeverity.Warning,
                    options.IncludeUnsupportedBlocksAsText
                        ? "Extension markup is represented as plain text in PowerPoint output."
                        : "Extension markup is omitted from PowerPoint output.",
                    block,
                    options.IncludeUnsupportedBlocksAsText
                        ? OfficeConversionLossKind.Approximation
                        : OfficeConversionLossKind.Omission));
            }
        }

        return diagnostics.AsReadOnly();
    }

    private static OfficeMarkupDiagnostic Omitted(OfficeMarkupBlock block) => new OfficeMarkupDiagnostic(
        OfficeMarkupDiagnosticSeverity.Warning,
        $"{block.Kind} markup is omitted from PowerPoint output.",
        block,
        OfficeConversionLossKind.Omission);
}
