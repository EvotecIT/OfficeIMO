namespace OfficeIMO.Epub;

/// <summary>Bounded ordered narration for one XHTML document. Audio is retained, not decoded.</summary>
public sealed class EpubMediaOverlay {
    /// <summary>One to 10,000 cues in playback order. Text targets must be distinct and in document order.</summary>
    public IReadOnlyList<EpubMediaOverlayCue> Cues { get; set; } = Array.Empty<EpubMediaOverlayCue>();
    /// <summary>Optional tree of cues and sequences instead of Cues. At most 10,000 total nodes and 32 nested sequences.</summary>
    public IReadOnlyList<EpubMediaOverlayNode> Nodes { get; set; } = Array.Empty<EpubMediaOverlayNode>();
    /// <summary>
    /// Caller-measured duration of each referenced audio resource, keyed by manifest id.
    /// These values bound clip offsets; they are not independently verified against encoded media.
    /// </summary>
    public IReadOnlyDictionary<string, TimeSpan> AudioDurations { get; set; } = new Dictionary<string, TimeSpan>();
}
