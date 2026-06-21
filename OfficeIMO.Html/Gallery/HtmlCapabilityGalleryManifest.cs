namespace OfficeIMO.Html;

/// <summary>
/// Structured manifest for an OfficeIMO HTML capability gallery result.
/// </summary>
public sealed class HtmlCapabilityGalleryManifest {
    /// <summary>
    /// Creates a gallery manifest.
    /// </summary>
    public HtmlCapabilityGalleryManifest(HtmlCapabilityGalleryResult result, HtmlConversionProfile profile, HtmlRoundTripScore roundTripScore, HtmlResourceManifest resourceManifest) {
        Result = result ?? throw new ArgumentNullException(nameof(result));
        Profile = profile;
        RoundTripScore = roundTripScore;
        ResourceManifest = resourceManifest;
    }

    /// <summary>Gallery result and artifacts.</summary>
    public HtmlCapabilityGalleryResult Result { get; }

    /// <summary>Shared conversion profile used by the scenario.</summary>
    public HtmlConversionProfile Profile { get; }

    /// <summary>Optional round-trip score.</summary>
    public HtmlRoundTripScore RoundTripScore { get; }

    /// <summary>Optional resource manifest.</summary>
    public HtmlResourceManifest ResourceManifest { get; }
}
