namespace OfficeIMO.Publisher.Internal;

internal sealed class PublisherParseContext {
    private readonly List<OfficeConversionFidelityDiagnostic> _diagnostics = new();
    private readonly HashSet<string> _diagnosticKeys = new(StringComparer.Ordinal);
    internal PublisherParseContext(PublisherReadOptions options, CancellationToken token) { Options = options; Token = token; }
    internal PublisherReadOptions Options { get; }
    internal CancellationToken Token { get; }
    internal int Records { get; private set; }
    internal long ImageBytes { get; private set; }
    private long _projectedElements, _projectedCharacters, _projectedGradientStops, _projectedPathCommands;
    private long _pathWorkItems;
    private long _imageProcessingBytes;
    private long _imageProcessingPixels;
    private long _textLayoutCharacters;
    private int _imageStoreEntries;
    private long _tableModelItems, _tableModelCharacters;
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
        Token.ThrowIfCancellationRequested();
        if (bytes > Options.Limits.MaxInputBytes - _imageProcessingBytes) throw new InvalidDataException("Publisher image processing byte limit exceeded.");
        _imageProcessingBytes += bytes;
    }
    internal long RemainingImageProcessingPixels => Options.MaximumImageProcessingPixels - _imageProcessingPixels;
    internal void AccountImageProcessingPixels(long pixels) {
        Token.ThrowIfCancellationRequested();
        if (pixels > RemainingImageProcessingPixels) throw new InvalidDataException("Publisher image processing pixel limit exceeded.");
        _imageProcessingPixels += pixels;
    }
    internal void AccountTextLayout(int characters) {
        Token.ThrowIfCancellationRequested();
        if (characters > Options.Limits.MaxTextCharacters - _textLayoutCharacters)
            throw new InvalidDataException("Publisher text layout character work limit exceeded.");
        _textLayoutCharacters += characters;
    }
    internal void AccountTableModel(int items) {
        Token.ThrowIfCancellationRequested();
        if (items > Options.Limits.MaxItems - _tableModelItems)
            throw new InvalidDataException("Publisher table model item limit exceeded.");
        _tableModelItems += items;
    }
    internal void AccountTableText(long characters) {
        Token.ThrowIfCancellationRequested();
        if (characters > Options.Limits.MaxTextCharacters - _tableModelCharacters)
            throw new InvalidDataException("Publisher table model text limit exceeded.");
        _tableModelCharacters += characters;
    }
    internal void AccountPathWork(int items) {
        Token.ThrowIfCancellationRequested();
        if (items > Options.Limits.MaxItems - _pathWorkItems)
            throw new InvalidDataException("Publisher custom path item work limit exceeded.");
        _pathWorkItems += items;
    }
    private void AccountPathProjection(int commands) {
        _projectedPathCommands += commands;
        if (_projectedPathCommands > Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher projected path command work limit exceeded.");
    }
    internal void AccountProjection(OfficeIMO.Drawing.OfficeDrawing drawing) {
        foreach (OfficeIMO.Drawing.OfficeDrawingElement element in drawing.Elements) {
            Token.ThrowIfCancellationRequested();
            if (++_projectedElements > Options.Limits.MaxItems) throw new InvalidDataException("Publisher projected drawing element limit exceeded.");
            if (element is OfficeIMO.Drawing.OfficeDrawingImage image) AccountImage(image.EncodedBytes.Length);
            if (element is OfficeIMO.Drawing.OfficeDrawingShape shape) {
                AccountPathProjection(shape.Shape.PathCommands.Count);
                _projectedGradientStops += (shape.Shape.FillGradient?.Stops.Count ?? 0) + (shape.Shape.FillRadialGradient?.Stops.Count ?? 0)
                    + (shape.Shape.StrokeGradient?.Stops.Count ?? 0) + (shape.Shape.StrokeRadialGradient?.Stops.Count ?? 0);
                if (_projectedGradientStops > Options.Limits.MaxItems)
                    throw new InvalidDataException("Publisher projected gradient stop work limit exceeded.");
            }
            if (element is OfficeIMO.Drawing.OfficeDrawingRichText text) {
                long characters = text.Paragraphs.Count != 0 ? text.Paragraphs.Sum(paragraph => paragraph.Runs.Sum(run => (long)run.Text.Length))
                    : text.Runs.Sum(run => (long)run.Text.Length);
                _projectedCharacters += characters;
                if (_projectedCharacters > Options.Limits.MaxTextCharacters) throw new InvalidDataException("Publisher projected text character limit exceeded.");
            }
            if (element is OfficeIMO.Drawing.OfficeDrawingGroup group) {
                AccountPathProjection(group.ClipPath.Commands.Count); AccountProjection(group.InnerDrawing);
            }
            if (element is OfficeIMO.Drawing.OfficeDrawingEffectGroup transformed) AccountProjection(transformed.InnerDrawing);
        }
    }
    internal void Add(string code, string message, OfficeConversionLossKind kind, string? location = null) {
        if (_diagnosticKeys.Add(code + "\n" + location))
            _diagnostics.Add(new OfficeConversionFidelityDiagnostic(code, message, kind, "OfficeIMO.Publisher", location));
    }
}
