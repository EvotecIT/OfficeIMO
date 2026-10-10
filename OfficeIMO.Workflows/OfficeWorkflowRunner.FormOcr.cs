using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Prepares a bounded, immutable review of visible values in existing form fields.</summary>
    /// <remarks>The caller owns the engine. This in-memory workflow captures limits/options, uses the canonical
    /// OCR provider admission, and neither accepts values nor publishes a file. Apply explicit review decisions
    /// through <see cref="PdfFormOcrReview.Apply"/> and the host's existing undo/publication owner.</remarks>
    public async Task<PdfFormOcrReview> PreparePdfFormOcrAsync(PdfDocument document, IOcrEngine engine,
        PdfOcrMergeOptions? options = null, OfficeWorkflowLimits? limits = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(engine);
        cancellationToken.ThrowIfCancellationRequested();
        var capturedLimits = (limits ?? new OfficeWorkflowLimits()).CloneAndValidate();
        var capturedOptions = options?.Clone() ?? new PdfOcrMergeOptions();
        byte[] bytes = document.ToBytes(cancellationToken);
        if (bytes.LongLength > capturedLimits.MaximumInputBytes)
            throw new InvalidOperationException($"The source PDF exceeds the configured {capturedLimits.MaximumInputBytes:N0}-byte input limit.");
        var source = PdfDocument.Load(bytes, document.ReadOptions);
        return await source.PrepareFormOcrAsync(engine, capturedOptions, cancellationToken).ConfigureAwait(false);
    }
}
