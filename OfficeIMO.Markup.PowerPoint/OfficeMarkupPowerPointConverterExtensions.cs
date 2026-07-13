using OfficeIMO.PowerPoint;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markup.PowerPoint;

/// <summary>Typed OfficeIMO Markup to PowerPoint conversion helpers.</summary>
public static class OfficeMarkupPowerPointConverterExtensions {
    /// <summary>Converts presentation-profile markup to an editable detached presentation.</summary>
    public static PowerPointPresentation ToPowerPointPresentation(
        this OfficeMarkupDocument document,
        MarkupToPowerPointOptions? options = null) {
        OfficeMarkupPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        if (result.Succeeded) return result.Value;
        result.Value.Dispose();
        return result.RequireValue();
    }

    /// <summary>Converts presentation-profile markup and returns the presentation with structured evidence.</summary>
    public static OfficeMarkupPowerPointConversionResult ToPowerPointPresentationResult(
        this OfficeMarkupDocument document,
        MarkupToPowerPointOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        MarkupToPowerPointOptions resolved = options ?? new MarkupToPowerPointOptions();
        IReadOnlyList<OfficeMarkupDiagnostic> diagnostics = CollectDiagnostics(document, resolved);
        return new OfficeMarkupPowerPointExporter().Build(document, resolved, diagnostics);
    }

    /// <summary>Converts presentation-profile markup and saves it as a PowerPoint presentation.</summary>
    public static OfficeMarkupPowerPointConversionReport SaveAsPowerPoint(
        this OfficeMarkupDocument document,
        string path,
        MarkupToPowerPointOptions? options = null) {
        OfficeMarkupPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) result.RequireValue().Save(path);
        return result.Report;
    }

    /// <summary>Converts presentation-profile markup and writes it to a caller-owned stream.</summary>
    public static OfficeMarkupPowerPointConversionReport SaveAsPowerPoint(
        this OfficeMarkupDocument document,
        Stream stream,
        MarkupToPowerPointOptions? options = null) {
        OfficeMarkupPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) result.RequireValue().Save(stream);
        return result.Report;
    }

    /// <summary>Converts presentation-profile markup and asynchronously saves it as a PowerPoint presentation.</summary>
    public static async Task<OfficeMarkupPowerPointConversionReport> SaveAsPowerPointAsync(
        this OfficeMarkupDocument document,
        string path,
        MarkupToPowerPointOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeMarkupPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) await result.RequireValue().SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Converts presentation-profile markup and asynchronously writes it to a caller-owned stream.</summary>
    public static async Task<OfficeMarkupPowerPointConversionReport> SaveAsPowerPointAsync(
        this OfficeMarkupDocument document,
        Stream stream,
        MarkupToPowerPointOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeMarkupPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) await result.RequireValue().SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    private static IReadOnlyList<OfficeMarkupDiagnostic> CollectDiagnostics(
        OfficeMarkupDocument document,
        MarkupToPowerPointOptions options) {
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
                    block));
            } else if (block is OfficeMarkupDiagramBlock && !options.RenderMermaidDiagrams) {
                diagnostics.Add(new OfficeMarkupDiagnostic(
                    OfficeMarkupDiagnosticSeverity.Warning,
                    "Diagram markup is represented as source text because diagram rendering is disabled.",
                    block));
            }
        }

        return diagnostics.AsReadOnly();
    }

    private static OfficeMarkupDiagnostic Omitted(OfficeMarkupBlock block) => new OfficeMarkupDiagnostic(
        OfficeMarkupDiagnosticSeverity.Warning,
        $"{block.Kind} markup is omitted from PowerPoint output.",
        block);
}
