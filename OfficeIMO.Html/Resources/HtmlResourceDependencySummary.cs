namespace OfficeIMO.Html;

/// <summary>
/// Resource dependency counts for one resource kind.
/// </summary>
public sealed class HtmlResourceDependencySummary {
    internal HtmlResourceDependencySummary(HtmlResourceKind kind, int totalCount, int allowedCount, int blockedCount) {
        Kind = kind;
        TotalCount = totalCount;
        AllowedCount = allowedCount;
        BlockedCount = blockedCount;
    }

    /// <summary>Resource kind represented by this summary.</summary>
    public HtmlResourceKind Kind { get; }

    /// <summary>Total discovered references of this kind.</summary>
    public int TotalCount { get; }

    /// <summary>References allowed by the configured policy.</summary>
    public int AllowedCount { get; }

    /// <summary>References blocked or rejected by the configured policy.</summary>
    public int BlockedCount { get; }
}
