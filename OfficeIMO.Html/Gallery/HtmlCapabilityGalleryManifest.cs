namespace OfficeIMO.Html;

/// <summary>
/// Structured manifest for an OfficeIMO HTML capability gallery result.
/// </summary>
public sealed class HtmlCapabilityGalleryManifest {
    private readonly IReadOnlyList<HtmlCapabilityGalleryExpectation> _expectations;
    private readonly IReadOnlyList<OfficeHtmlConversionProfile> _officeProfiles;

    /// <summary>
    /// Creates a gallery manifest.
    /// </summary>
    public HtmlCapabilityGalleryManifest(
        HtmlCapabilityGalleryResult result,
        HtmlConversionProfile profile,
        HtmlRoundTripScore roundTripScore,
        HtmlResourceManifest resourceManifest)
        : this(result, profile, roundTripScore, resourceManifest, Array.Empty<HtmlCapabilityGalleryExpectation>()) {
    }

    /// <summary>
    /// Creates a gallery manifest with explicit proof expectations.
    /// </summary>
    public HtmlCapabilityGalleryManifest(
        HtmlCapabilityGalleryResult result,
        HtmlConversionProfile profile,
        HtmlRoundTripScore roundTripScore,
        HtmlResourceManifest resourceManifest,
        IEnumerable<HtmlCapabilityGalleryExpectation> expectations)
        : this(result, profile, roundTripScore, resourceManifest, expectations, Array.Empty<OfficeHtmlConversionProfile>()) {
    }

    /// <summary>
    /// Creates a gallery manifest with explicit proof expectations and source-specific Office-to-HTML profile contracts.
    /// </summary>
    public HtmlCapabilityGalleryManifest(
        HtmlCapabilityGalleryResult result,
        HtmlConversionProfile profile,
        HtmlRoundTripScore roundTripScore,
        HtmlResourceManifest resourceManifest,
        IEnumerable<HtmlCapabilityGalleryExpectation> expectations,
        IEnumerable<OfficeHtmlConversionProfile> officeProfiles) {
        Result = result ?? throw new ArgumentNullException(nameof(result));
        Profile = profile;
        RoundTripScore = roundTripScore;
        ResourceManifest = resourceManifest;
        _expectations = ToReadOnlyList(expectations);
        _officeProfiles = ToOfficeProfileList(officeProfiles);
    }

    /// <summary>Gallery result and artifacts.</summary>
    public HtmlCapabilityGalleryResult Result { get; }

    /// <summary>Shared conversion profile used by the scenario.</summary>
    public HtmlConversionProfile Profile { get; }

    /// <summary>Optional round-trip score.</summary>
    public HtmlRoundTripScore RoundTripScore { get; }

    /// <summary>Optional resource manifest.</summary>
    public HtmlResourceManifest ResourceManifest { get; }

    /// <summary>Expected proof outcomes for features, simplifications, blocked resources, diagnostics, and visual or logical evidence.</summary>
    public IReadOnlyList<HtmlCapabilityGalleryExpectation> Expectations => _expectations;

    /// <summary>Source-specific Office-to-HTML profile contracts proven by this gallery scenario.</summary>
    public IReadOnlyList<OfficeHtmlConversionProfile> OfficeProfiles => _officeProfiles;

    private static IReadOnlyList<HtmlCapabilityGalleryExpectation> ToReadOnlyList(IEnumerable<HtmlCapabilityGalleryExpectation> expectations) {
        if (expectations == null) {
            throw new ArgumentNullException(nameof(expectations));
        }

        return expectations.Where(expectation => expectation != null).ToList().AsReadOnly();
    }

    private static IReadOnlyList<OfficeHtmlConversionProfile> ToOfficeProfileList(IEnumerable<OfficeHtmlConversionProfile> officeProfiles) {
        if (officeProfiles == null) {
            throw new ArgumentNullException(nameof(officeProfiles));
        }

        return officeProfiles.Distinct().ToList().AsReadOnly();
    }
}
