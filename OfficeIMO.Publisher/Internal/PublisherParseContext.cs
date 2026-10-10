namespace OfficeIMO.Publisher.Internal;

internal sealed class PublisherParseContext {
    private readonly List<OfficeConversionFidelityDiagnostic> _diagnostics = new();
    private readonly HashSet<string> _diagnosticKeys = new(StringComparer.Ordinal);
    internal PublisherParseContext(PublisherReadOptions options, CancellationToken token) { Options = options; Token = token; }
    internal PublisherReadOptions Options { get; }
    internal CancellationToken Token { get; }
    internal int Records { get; private set; }
    internal long ImageBytes { get; private set; }
    private long _projectedElements, _projectedCharacters;
    private long _imageProcessingBytes;
    private long _textLayoutCharacters;
    private int _imageStoreEntries;
    internal IReadOnlyList<OfficeConversionFidelityDiagnostic> Diagnostics => _diagnostics;
    internal void Record() {
        Token.ThrowIfCancellationRequested();
        if (Records >= Options.Limits.MaxRecords) throw new InvalidDataException("Publisher native record limit exceeded.");
        Records++;
    }
    internal void CheckDepth(int depth) {
        Token.ThrowIfCancellationRequested();
        if (depth > Options.MaximumNestingDepth) throw new InvalidDataException("Publisher nesting limit exceeded.");
    }
    internal void AccountImage(int bytes) {
        if (bytes > Options.MaximumTotalImageBytes - ImageBytes) throw new InvalidDataException("Publisher total image byte limit exceeded.");
        ImageBytes += bytes;
    }
    internal void AccountImageStoreEntry() {
        Token.ThrowIfCancellationRequested();
        if (++_imageStoreEntries > Options.Limits.MaxItems) throw new InvalidDataException("Publisher image store entry limit exceeded.");
    }
    internal void AccountImageProcessing(int bytes) {
        if (bytes > Options.Limits.MaxInputBytes - _imageProcessingBytes) throw new InvalidDataException("Publisher image processing byte limit exceeded.");
        _imageProcessingBytes += bytes;
    }
    internal void AccountTextLayout(int characters) {
        Token.ThrowIfCancellationRequested();
        if (characters > Options.Limits.MaxTextCharacters - _textLayoutCharacters)
            throw new InvalidDataException("Publisher text layout character work limit exceeded.");
        _textLayoutCharacters += characters;
    }
    internal void AccountProjection(OfficeIMO.Drawing.OfficeDrawing drawing) {
        foreach (OfficeIMO.Drawing.OfficeDrawingElement element in drawing.Elements) {
            Token.ThrowIfCancellationRequested();
            if (++_projectedElements > Options.Limits.MaxItems) throw new InvalidDataException("Publisher projected drawing element limit exceeded.");
            if (element is OfficeIMO.Drawing.OfficeDrawingImage image) AccountImage(image.EncodedBytes.Length);
            if (element is OfficeIMO.Drawing.OfficeDrawingRichText text) {
                long characters = text.Paragraphs.Count != 0 ? text.Paragraphs.Sum(paragraph => paragraph.Runs.Sum(run => (long)run.Text.Length))
                    : text.Runs.Sum(run => (long)run.Text.Length);
                _projectedCharacters += characters;
                if (_projectedCharacters > Options.Limits.MaxTextCharacters) throw new InvalidDataException("Publisher projected text character limit exceeded.");
            }
            if (element is OfficeIMO.Drawing.OfficeDrawingGroup group) AccountProjection(group.InnerDrawing);
        }
    }
    internal void Add(string code, string message, OfficeConversionLossKind kind, string? location = null) {
        if (_diagnosticKeys.Add(code + "\n" + location))
            _diagnostics.Add(new OfficeConversionFidelityDiagnostic(code, message, kind, "OfficeIMO.Publisher", location));
    }
}
