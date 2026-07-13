using OfficeIMO.Word;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markup.Word;

/// <summary>Typed OfficeIMO Markup to Word conversion helpers.</summary>
public static class OfficeMarkupWordConverterExtensions {
    /// <summary>Converts document-profile markup to an editable detached Word document.</summary>
    public static WordDocument ToWordDocument(this OfficeMarkupDocument document, MarkupToWordOptions? options = null) {
        OfficeMarkupConversionResult<WordDocument> result = document.ToWordDocumentResult(options);
        if (result.Succeeded) return result.Value;
        result.Value.Dispose();
        return result.RequireValue();
    }

    /// <summary>Converts document-profile markup and returns the Word document with structured diagnostics.</summary>
    public static OfficeMarkupConversionResult<WordDocument> ToWordDocumentResult(
        this OfficeMarkupDocument document,
        MarkupToWordOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        MarkupToWordOptions resolved = options ?? new MarkupToWordOptions();
        IReadOnlyList<OfficeMarkupDiagnostic> diagnostics = CollectDiagnostics(document, resolved);
        WordDocument value = new OfficeMarkupWordExporter().Convert(document, resolved);
        return new OfficeMarkupConversionResult<WordDocument>(value, diagnostics);
    }

    /// <summary>Converts document-profile markup and saves it as a Word document.</summary>
    public static OfficeMarkupConversionReport SaveAsWord(
        this OfficeMarkupDocument document,
        string path,
        MarkupToWordOptions? options = null,
        WordSaveOptions? saveOptions = null) {
        OfficeMarkupConversionResult<WordDocument> result = document.ToWordDocumentResult(options);
        using (result.Value) result.RequireValue().Save(path, saveOptions);
        return result.Report;
    }

    /// <summary>Converts document-profile markup and writes it to a caller-owned stream.</summary>
    public static OfficeMarkupConversionReport SaveAsWord(
        this OfficeMarkupDocument document,
        Stream stream,
        MarkupToWordOptions? options = null,
        WordSaveOptions? saveOptions = null) {
        OfficeMarkupConversionResult<WordDocument> result = document.ToWordDocumentResult(options);
        using (result.Value) result.RequireValue().Save(stream, saveOptions);
        return result.Report;
    }

    /// <summary>Converts document-profile markup and asynchronously saves it as a Word document.</summary>
    public static async Task<OfficeMarkupConversionReport> SaveAsWordAsync(
        this OfficeMarkupDocument document,
        string path,
        MarkupToWordOptions? options = null,
        WordSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeMarkupConversionResult<WordDocument> result = document.ToWordDocumentResult(options);
        using (result.Value) await result.RequireValue().SaveAsync(path, saveOptions, cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Converts document-profile markup and asynchronously writes it to a caller-owned stream.</summary>
    public static async Task<OfficeMarkupConversionReport> SaveAsWordAsync(
        this OfficeMarkupDocument document,
        Stream stream,
        MarkupToWordOptions? options = null,
        WordSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeMarkupConversionResult<WordDocument> result = document.ToWordDocumentResult(options);
        using (result.Value) await result.RequireValue().SaveAsync(stream, saveOptions, cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    private static IReadOnlyList<OfficeMarkupDiagnostic> CollectDiagnostics(
        OfficeMarkupDocument document,
        MarkupToWordOptions options) {
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
                    block));
            }
        }

        return diagnostics.AsReadOnly();
    }
}
