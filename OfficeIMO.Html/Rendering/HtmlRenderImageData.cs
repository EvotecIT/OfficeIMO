namespace OfficeIMO.Html;

internal sealed class HtmlRenderImageData {
    private readonly byte[] _bytes;

    internal HtmlRenderImageData(byte[] bytes) {
        if (bytes == null || bytes.Length == 0) {
            throw new ArgumentException("Rendered images require encoded bytes.", nameof(bytes));
        }

        _bytes = (byte[])bytes.Clone();
    }

    internal byte[] EncodedBytes => _bytes;

    internal byte[] Snapshot() => (byte[])_bytes.Clone();
}
