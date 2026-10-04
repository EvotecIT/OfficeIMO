namespace OfficeIMO.AsciiDoc;

/// <summary>A source block and the attribute snapshot effective at its position.</summary>
public sealed class AsciiDocBlockContext {
    internal AsciiDocBlockContext(AsciiDocBlock block, AsciiDocDocumentAttributes attributes) {
        Block = block;
        Attributes = attributes;
    }
    /// <summary>Source-backed semantic block.</summary>
    public AsciiDocBlock Block { get; }
    /// <summary>Attributes after preceding source assignments, including this block when it is an assignment.</summary>
    public AsciiDocDocumentAttributes Attributes { get; }
}
