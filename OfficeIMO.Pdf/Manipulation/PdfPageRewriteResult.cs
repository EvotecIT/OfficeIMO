namespace OfficeIMO.Pdf;

/// <summary>A page rewrite and its original-to-output object-number mapping.</summary>
public sealed class PdfPageRewriteResult {
    private readonly byte[] _bytes;
    private readonly PdfLoadOptions? _readOptions;

    internal PdfPageRewriteResult(byte[] bytes, IReadOnlyDictionary<int, int> objectNumberMap, PdfLoadOptions? readOptions) {
        _bytes = (byte[])bytes.Clone();
        _readOptions = PdfLoadOptions.WithMinimumInputBytes(readOptions, _bytes.LongLength);
        ObjectNumberMap = new System.Collections.ObjectModel.ReadOnlyDictionary<int, int>(objectNumberMap.ToDictionary(pair => pair.Key, pair => pair.Value));
    }

    /// <summary>Rewritten PDF bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Mapping for retained original objects. Newly generated objects have no original key.</summary>
    public IReadOnlyDictionary<int, int> ObjectNumberMap { get; }

    /// <summary>Opens the rewritten document with the source read options.</summary>
    public PdfDocument ToDocument() => PdfDocument.Load(_bytes, _readOptions);
}
