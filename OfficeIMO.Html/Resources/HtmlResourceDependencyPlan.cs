namespace OfficeIMO.Html;

/// <summary>
/// Adapter-friendly resource dependency plan derived from an HTML resource manifest.
/// </summary>
public sealed class HtmlResourceDependencyPlan {
    internal HtmlResourceDependencyPlan(
        IReadOnlyList<HtmlResourceReference> allowedResources,
        IReadOnlyList<HtmlResourceReference> blockedResources,
        IReadOnlyList<HtmlResourceDependencySummary> summaries) {
        AllowedResources = allowedResources ?? throw new ArgumentNullException(nameof(allowedResources));
        BlockedResources = blockedResources ?? throw new ArgumentNullException(nameof(blockedResources));
        Summaries = summaries ?? throw new ArgumentNullException(nameof(summaries));
    }

    /// <summary>Allowed resources in source order.</summary>
    public IReadOnlyList<HtmlResourceReference> AllowedResources { get; }

    /// <summary>Blocked resources in source order.</summary>
    public IReadOnlyList<HtmlResourceReference> BlockedResources { get; }

    /// <summary>Resource counts grouped by kind.</summary>
    public IReadOnlyList<HtmlResourceDependencySummary> Summaries { get; }

    /// <summary>Whether the plan contains blocked resources.</summary>
    public bool HasBlockedResources => BlockedResources.Count > 0;

    /// <summary>Gets a summary for one kind, returning zero counts when the kind is absent.</summary>
    public HtmlResourceDependencySummary GetSummary(HtmlResourceKind kind) {
        foreach (HtmlResourceDependencySummary summary in Summaries) {
            if (summary.Kind == kind) {
                return summary;
            }
        }

        return new HtmlResourceDependencySummary(kind, 0, 0, 0);
    }
}
