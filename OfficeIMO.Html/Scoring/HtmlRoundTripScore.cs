namespace OfficeIMO.Html;

/// <summary>
/// Structural score comparing source HTML with generated or round-tripped HTML.
/// </summary>
public sealed class HtmlRoundTripScore {
    internal HtmlRoundTripScore(
        double score,
        int sourceNodeCount,
        int targetNodeCount,
        int matchedFeatureCount,
        int comparedFeatureCount,
        IReadOnlyDictionary<string, double> metrics,
        IReadOnlyDictionary<string, double>? dimensions = null,
        bool artifactReloadVerified = false,
        string? artifactKind = null) {
        Score = score;
        SourceNodeCount = sourceNodeCount;
        TargetNodeCount = targetNodeCount;
        MatchedFeatureCount = matchedFeatureCount;
        ComparedFeatureCount = comparedFeatureCount;
        Metrics = metrics;
        Dimensions = dimensions ?? metrics;
        ArtifactReloadVerified = artifactReloadVerified;
        ArtifactKind = artifactKind;
    }

    /// <summary>Current fidelity-score schema version.</summary>
    public const int CurrentSchemaVersion = 2;

    /// <summary>Fidelity-score schema version used for this result.</summary>
    public int SchemaVersion => CurrentSchemaVersion;

    /// <summary>Overall score from 0 to 1.</summary>
    public double Score { get; }

    /// <summary>Logical node count in the source document.</summary>
    public int SourceNodeCount { get; }

    /// <summary>Logical node count in the target document.</summary>
    public int TargetNodeCount { get; }

    /// <summary>Feature buckets with exact or near matches.</summary>
    public int MatchedFeatureCount { get; }

    /// <summary>Feature buckets compared by the scorer.</summary>
    public int ComparedFeatureCount { get; }

    /// <summary>Named scoring metrics from 0 to 1.</summary>
    public IReadOnlyDictionary<string, double> Metrics { get; }

    /// <summary>
    /// Top-level v2 fidelity dimensions. Available dimensions are structure, text, styles,
    /// resources, annotations, formulas, charts, geometry, and artifact reload.
    /// Dimensions absent from both inputs are omitted rather than treated as perfect.
    /// </summary>
    public IReadOnlyDictionary<string, double> Dimensions { get; }

    /// <summary>Whether the score includes evidence produced after reopening the native artifact.</summary>
    public bool ArtifactReloadVerified { get; }

    /// <summary>Caller-supplied artifact kind for reload evidence, such as DOCX, XLSX, or PPTX.</summary>
    public string? ArtifactKind { get; }
}
