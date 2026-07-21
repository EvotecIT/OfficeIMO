namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentStamper {
    /// <summary>
    /// Stamps arbitrary visual <see cref="PdfPageCanvas"/> content onto selected existing pages.
    /// Annotations and form fields remain separate editor concerns and are rejected by this visual-only operation.
    /// </summary>
    public PdfDocument Content(
        Action<PdfPageCanvas, PdfStampPageContext> build,
        PdfCanvasStampOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(build, nameof(build));
        return _document.ApplyMutation(input => PdfStamper.StampCanvas(input, build, options, readOptions ?? _document.ReadOptions));
    }

    /// <summary>Stamps the same arbitrary visual canvas content onto selected existing pages.</summary>
    public PdfDocument Content(
        Action<PdfPageCanvas> build,
        PdfCanvasStampOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(build, nameof(build));
        return Content((canvas, _) => build(canvas), options, readOptions);
    }

    /// <summary>Attempts to stamp arbitrary visual canvas content and returns diagnostics when blocked or failed.</summary>
    public PdfOperationResult<PdfDocument> TryContent(
        Action<PdfPageCanvas, PdfStampPageContext> build,
        PdfCanvasStampOptions? stampOptions = null,
        PdfReadOptions? options = null) {
        Guard.NotNull(build, nameof(build));
        return _document.TryMutationOperation(
            "Stamp visual canvas content",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ModifyPageContent,
            _ => Content(build, stampOptions, options),
            options: options);
    }
}
