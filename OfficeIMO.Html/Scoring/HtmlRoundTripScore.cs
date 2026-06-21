namespace OfficeIMO.Html;

/// <summary>
/// Structural score comparing source HTML with generated or round-tripped HTML.
/// </summary>
public sealed class HtmlRoundTripScore {
    internal HtmlRoundTripScore(double score, int sourceNodeCount, int targetNodeCount, int matchedFeatureCount, int comparedFeatureCount, IReadOnlyDictionary<string, double> metrics) {
        Score = score;
        SourceNodeCount = sourceNodeCount;
        TargetNodeCount = targetNodeCount;
        MatchedFeatureCount = matchedFeatureCount;
        ComparedFeatureCount = comparedFeatureCount;
        Metrics = metrics;
    }

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
}
