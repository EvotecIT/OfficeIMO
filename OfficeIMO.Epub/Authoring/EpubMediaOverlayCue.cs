namespace OfficeIMO.Epub;

/// <summary>One narrated XHTML element and an explicit interval in a retained audio resource.</summary>
public sealed class EpubMediaOverlayCue : EpubMediaOverlayNode {
    /// <summary>Creates a cue. Targets and timings are checked when the overlay is added.</summary>
    public EpubMediaOverlayCue(string elementId, string audioManifestId, TimeSpan clipBegin, TimeSpan clipEnd) : base(elementId) {
        AudioManifestId = audioManifestId ?? throw new ArgumentNullException(nameof(audioManifestId));
        ClipBegin = clipBegin; ClipEnd = clipEnd;
    }
    /// <summary>Manifest id of a local, unencrypted audio/mpeg or audio/mp4 resource.</summary>
    public string AudioManifestId { get; }
    /// <summary>Inclusive nonnegative start offset.</summary>
    public TimeSpan ClipBegin { get; }
    /// <summary>Exclusive end offset, greater than the start and no later than the declared audio duration.</summary>
    public TimeSpan ClipEnd { get; }
}
