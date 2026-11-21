namespace OfficeIMO.Pdf;

internal sealed partial class PdfReadDocument {
    private PdfPageExtractor.CatalogRewriteState? _catalogRewriteState;
    private RewriteGraphValidation? _rewriteGraphValidation;

    /// <summary>
    /// Returns the catalog rewrite state shared by every page-transfer operation on this parse.
    /// Encrypted sources build a fresh state each time so no operation shares another's security handler.
    /// </summary>
    internal PdfPageExtractor.CatalogRewriteState GetCatalogRewriteState(CancellationToken cancellationToken) {
        if (Security.HasEncryption) {
            return PdfPageExtractor.ExtractCatalogRewriteState(_objects, _trailerRaw, this, cancellationToken);
        }

        PdfPageExtractor.CatalogRewriteState? state = Volatile.Read(ref _catalogRewriteState);
        if (state is not null) {
            return state;
        }

        // The retained state outlives this caller, so its deferred indexes must not observe the caller's token.
        state = PdfPageExtractor.ExtractCatalogRewriteState(_objects, _trailerRaw, cancellationToken, CancellationToken.None);
        return Interlocked.CompareExchange(ref _catalogRewriteState, state, null) ?? state;
    }

    /// <summary>
    /// Returns the retained outcome of rewrite object-graph validation, running <paramref name="validate"/> once.
    /// The validator returns the blocker message, or null when the graph can be rewritten.
    /// </summary>
    internal string? GetRewriteObjectGraphFailure(Func<PdfReadDocument, CancellationToken, string?> validate, CancellationToken cancellationToken) {
        RewriteGraphValidation? validation = Volatile.Read(ref _rewriteGraphValidation);
        if (validation is null) {
            validation = new RewriteGraphValidation(validate(this, cancellationToken));
            validation = Interlocked.CompareExchange(ref _rewriteGraphValidation, validation, null) ?? validation;
        }

        return validation.Failure;
    }

    private sealed class RewriteGraphValidation {
        internal RewriteGraphValidation(string? failure) => Failure = failure;

        internal string? Failure { get; }
    }
}
