namespace OfficeIMO.Html;

/// <summary>
/// Result manifest for one HTML capability-gallery scenario.
/// </summary>
public sealed class HtmlCapabilityGalleryResult {
    private readonly List<HtmlCapabilityGalleryArtifact> _artifacts = new List<HtmlCapabilityGalleryArtifact>();
    private readonly IReadOnlyList<HtmlCapabilityGalleryArtifact> _readOnlyArtifacts;

    /// <summary>
    /// Creates a mutable capability-gallery result builder for compatibility with existing producers.
    /// A manifest takes an immutable snapshot when it is constructed.
    /// </summary>
    /// <param name="scenario">Scenario proven by the result.</param>
    public HtmlCapabilityGalleryResult(HtmlCapabilityGalleryScenario scenario) {
        Scenario = scenario ?? throw new ArgumentNullException(nameof(scenario));
        _readOnlyArtifacts = _artifacts.AsReadOnly();
        Diagnostics = new HtmlDiagnosticReport();
    }

    /// <summary>
    /// Creates a capability-gallery result.
    /// </summary>
    /// <param name="scenario">Scenario proven by the result.</param>
    /// <param name="artifacts">Artifacts emitted for the scenario. The sequence is copied.</param>
    /// <param name="diagnostics">Diagnostics captured while generating the scenario. The sequence is copied.</param>
    public HtmlCapabilityGalleryResult(
        HtmlCapabilityGalleryScenario scenario,
        IEnumerable<HtmlCapabilityGalleryArtifact> artifacts,
        IEnumerable<HtmlDiagnostic> diagnostics)
        : this(scenario) {
        if (artifacts == null) throw new ArgumentNullException(nameof(artifacts));
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));

        foreach (HtmlCapabilityGalleryArtifact artifact in artifacts) {
            AddArtifact(artifact ?? throw new ArgumentException("Gallery artifacts cannot contain null values.", nameof(artifacts)));
        }

        foreach (HtmlDiagnostic diagnostic in diagnostics) {
            Diagnostics.Add(diagnostic ?? throw new ArgumentException("Gallery diagnostics cannot contain null values.", nameof(diagnostics)));
        }

        Diagnostics = Diagnostics.Snapshot();
        IsReadOnly = true;
    }

    /// <summary>
    /// Scenario proven by the result.
    /// </summary>
    public HtmlCapabilityGalleryScenario Scenario { get; }

    /// <summary>
    /// Artifacts emitted for the scenario.
    /// </summary>
    public IReadOnlyList<HtmlCapabilityGalleryArtifact> Artifacts => _readOnlyArtifacts;

    /// <summary>
    /// Shared diagnostics captured while generating the scenario artifacts.
    /// </summary>
    public HtmlDiagnosticReport Diagnostics { get; private set; }

    /// <summary>True when this result is a finalized immutable snapshot.</summary>
    public bool IsReadOnly { get; private set; }

    /// <summary>Adds an artifact to a mutable result builder.</summary>
    public void AddArtifact(HtmlCapabilityGalleryArtifact artifact) {
        if (artifact == null) throw new ArgumentNullException(nameof(artifact));
        EnsureMutable();
        _artifacts.Add(artifact);
    }

    internal HtmlCapabilityGalleryResult Snapshot() =>
        new HtmlCapabilityGalleryResult(Scenario, _artifacts, Diagnostics);

    private void EnsureMutable() {
        if (IsReadOnly) {
            throw new InvalidOperationException("The capability-gallery result is an immutable manifest snapshot.");
        }
    }
}
