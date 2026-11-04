namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal sealed partial class PageContentBudget {
        private long _projectedStrokeDashEntries;

        internal void ChargeStrokeDashProjection(int count) {
            CancellationToken.ThrowIfCancellationRequested();
            _projectedStrokeDashEntries += count;
            if (_projectedStrokeDashEntries > _page._limits.MaxContentOperands) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperands,
                    _page._limits.MaxContentOperands, _projectedStrokeDashEntries);
            }
        }
    }
}
