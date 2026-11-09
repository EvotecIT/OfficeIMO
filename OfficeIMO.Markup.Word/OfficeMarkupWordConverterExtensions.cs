using OfficeIMO.Word;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Markup.Word;

/// <summary>Converts typed OfficeIMO Markup into editable Word documents.</summary>
public static class OfficeMarkupWordConverterExtensions {
    /// <summary>Converts markup into a detached document owned by the caller.</summary>
    public static WordDocument ToWord(this OfficeMarkupDocument document, MarkupWordExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToWord(options, out _, cancellationToken);

    /// <summary>Converts markup and retains mapping and native validation evidence.</summary>
    public static WordDocument ToWord(this OfficeMarkupDocument document, MarkupWordExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        cancellationToken.ThrowIfCancellationRequested();
        MarkupWordExportOptions resolved = options ?? new MarkupWordExportOptions();
        IReadOnlyList<OfficeMarkupDiagnostic> diagnostics = CollectDiagnostics(document, resolved);
        WordDocument value = new OfficeMarkupWordExporter().Convert(document, resolved);
        var result = new OfficeMarkupConversionResult<WordDocument>(value, diagnostics);
        report = result.Report;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (!result.Succeeded) throw new OfficeConversionException("Markup conversion failed validation.", report);
            if (resolved.FailOnLoss) report.ThrowIfLoss();
            return result.Value;
        } catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Converts markup asynchronously and returns the document with its completed report.</summary>
    public static Task<OfficeConversionResult<WordDocument>> ToWordAsync(this OfficeMarkupDocument document,
        MarkupWordExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToWord(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes converted Word output and returns its report.</summary>
    public static OfficeConversionReport SaveAsWord(
        this OfficeMarkupDocument document, string path, MarkupWordExportOptions? options = null, WordSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(path, saveOptions));
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Writes converted Word output and returns its report.</summary>
    public static OfficeConversionReport SaveAsWord(
        this OfficeMarkupDocument document, Stream stream, MarkupWordExportOptions? options = null, WordSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(stream, saveOptions));
        return report;
    }

    /// <summary>Asynchronously writes converted Word output and returns its report.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(
        this OfficeMarkupDocument document, string path, MarkupWordExportOptions? options = null, WordSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, saveOptions, cancellationToken)).ConfigureAwait(false);
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Asynchronously writes converted Word output and returns its report.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(
        this OfficeMarkupDocument document, Stream stream, MarkupWordExportOptions? options = null, WordSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, saveOptions, cancellationToken)).ConfigureAwait(false);
        return report;
    }

    private static IReadOnlyList<OfficeMarkupDiagnostic> CollectDiagnostics(
        OfficeMarkupDocument document,
        MarkupWordExportOptions options) {
        var diagnostics = new List<OfficeMarkupDiagnostic>(OfficeMarkupValidator.Validate(document));
        foreach (OfficeMarkupBlock block in document.DescendantsAndSelf()) {
            if (block is OfficeMarkupDiagramBlock) {
                diagnostics.Add(new OfficeMarkupDiagnostic(
                    OfficeMarkupDiagnosticSeverity.Warning,
                    "Diagram markup is represented as labeled source text in Word output.",
                    block));
            } else if (block is OfficeMarkupExtensionBlock || block is OfficeMarkupRawMarkdownBlock) {
                diagnostics.Add(new OfficeMarkupDiagnostic(
                    OfficeMarkupDiagnosticSeverity.Warning,
                    options.IncludeUnsupportedBlocksAsText
                        ? $"{block.Kind} markup is represented as plain text in Word output."
                        : $"{block.Kind} markup is omitted from Word output.",
                    block,
                    options.IncludeUnsupportedBlocksAsText
                        ? OfficeConversionLossKind.Approximation
                        : OfficeConversionLossKind.Omission));
            }
        }

        return diagnostics.AsReadOnly();
    }
}
