using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>Bakes selected pages' effective quarter-turn rotation into their content and geometry while preserving appearance.</summary>
    /// <remarks>This rewrites the active object graph. Signed or certified input retains its rewrite restrictions.</remarks>
    public PdfDocument NormalizeRotation(PdfPageSelection? selection = null, CancellationToken cancellationToken = default) =>
        _document.ApplyMutation(input => PdfPageEditor.NormalizePageRotation(input, selection, _document.ReadOptions, cancellationToken),
            cancellationToken: cancellationToken);

    /// <summary>Normalizes saved rotation and reports the rewrite route and restrictions.</summary>
    /// <exception cref="OperationCanceledException">The caller requests cancellation.</exception>
    public PdfOperationResult<PdfDocument> NormalizeRotationResult(PdfPageSelection? selection = null,
        PdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Normalize page rotation", PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.NormalizePageRotation, () => _document.ApplyMutation(
                input => PdfPageEditor.NormalizePageRotation(input, selection, effective, cancellationToken), effective,
                operationName: nameof(NormalizeRotation), cancellationToken: cancellationToken),
            effective, cancellationToken: cancellationToken);
    }
}
