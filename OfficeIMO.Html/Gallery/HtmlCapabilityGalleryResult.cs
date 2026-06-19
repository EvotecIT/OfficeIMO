namespace OfficeIMO.Html;

/// <summary>
/// Result manifest for one HTML capability-gallery scenario.
/// </summary>
public sealed class HtmlCapabilityGalleryResult {
    private readonly List<HtmlCapabilityGalleryArtifact> _artifacts = new List<HtmlCapabilityGalleryArtifact>();

    /// <summary>
    /// Creates a capability-gallery result.
    /// </summary>
    /// <param name="scenario">Scenario proven by the result.</param>
    public HtmlCapabilityGalleryResult(HtmlCapabilityGalleryScenario scenario) {
        Scenario = scenario ?? throw new ArgumentNullException(nameof(scenario));
    }

    /// <summary>
    /// Scenario proven by the result.
    /// </summary>
    public HtmlCapabilityGalleryScenario Scenario { get; }

    /// <summary>
    /// Artifacts emitted for the scenario.
    /// </summary>
    public IReadOnlyList<HtmlCapabilityGalleryArtifact> Artifacts => _artifacts;

    /// <summary>
    /// Shared diagnostics captured while generating the scenario artifacts.
    /// </summary>
    public HtmlDiagnosticReport Diagnostics { get; } = new HtmlDiagnosticReport();

    /// <summary>
    /// Adds an artifact to the result.
    /// </summary>
    /// <param name="artifact">Artifact descriptor to add.</param>
    public void AddArtifact(HtmlCapabilityGalleryArtifact artifact) {
        if (artifact == null) {
            throw new ArgumentNullException(nameof(artifact));
        }

        _artifacts.Add(artifact);
    }
}
