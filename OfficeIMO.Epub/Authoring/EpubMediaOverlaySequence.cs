namespace OfficeIMO.Epub;

/// <summary>A nested narration sequence whose children target descendants of one XHTML element.</summary>
public sealed class EpubMediaOverlaySequence : EpubMediaOverlayNode {
    /// <summary>Creates a sequence with a snapshot of one to 10,000 ordered child nodes.</summary>
    public EpubMediaOverlaySequence(string elementId, IReadOnlyList<EpubMediaOverlayNode> children) : base(elementId) {
        if (children == null) throw new ArgumentNullException(nameof(children));
        if (children.Count < 1 || children.Count > 10000) throw new ArgumentException("A sequence requires one to 10,000 child nodes.", nameof(children));
        Children = Array.AsReadOnly(children.ToArray());
    }
    /// <summary>Child cues and sequences in playback order. The writer permits at most 32 nested sequences.</summary>
    public IReadOnlyList<EpubMediaOverlayNode> Children { get; }
}
