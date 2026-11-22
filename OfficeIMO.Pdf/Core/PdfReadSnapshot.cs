namespace OfficeIMO.Pdf;

/// <summary>Pairs the canonical parse with its encoded source when an operation needs both.</summary>
internal readonly struct PdfReadSnapshot {
    private readonly PdfDocumentSource? _source;
    private readonly ReadOnlyMemory<byte> _generatedBytes;
    private readonly CancellationToken _cancellationToken;

    internal PdfReadSnapshot(PdfDocumentSource source, PdfReadDocument document, PdfLoadOptions options,
        CancellationToken cancellationToken) {
        _source = source;
        _cancellationToken = cancellationToken;
        Document = document;
        Options = options;
    }

    internal PdfReadSnapshot(ReadOnlyMemory<byte> generatedBytes, PdfReadDocument document, PdfLoadOptions options) {
        _generatedBytes = generatedBytes;
        Document = document;
        Options = options;
    }

    internal PdfReadDocument Document { get; }
    internal PdfLoadOptions Options { get; }
    internal ReadOnlyMemory<byte> Bytes => _source?.GetMemory(_cancellationToken) ?? _generatedBytes;
    internal ReadOnlyMemory<byte> InspectionBytes =>
        _source?.GetInspectionBytes(Document, _cancellationToken) ?? _generatedBytes;
}
