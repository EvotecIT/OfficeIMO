namespace OfficeIMO.Pdf;

internal sealed class RenderedType3TextTracker {
    private readonly HashSet<double> _paintOrders = new HashSet<double>();
    private readonly HashSet<PdfContentOrderKey> _contentOrderKeys = new HashSet<PdfContentOrderKey>();

    internal void Add(double paintOrder, PdfContentOrderKey? contentOrderKey) {
        if (contentOrderKey != null) {
            _contentOrderKeys.Add(contentOrderKey);
        } else {
            _paintOrders.Add(paintOrder);
        }
    }

    internal bool Contains(double paintOrder, PdfContentOrderKey? contentOrderKey) =>
        contentOrderKey != null
            ? _contentOrderKeys.Contains(contentOrderKey)
            : _paintOrders.Contains(paintOrder);
}
