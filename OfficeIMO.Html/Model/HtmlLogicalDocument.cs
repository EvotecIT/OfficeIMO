namespace OfficeIMO.Html;

/// <summary>
/// Shared logical representation of parsed HTML before adapter-specific conversion.
/// </summary>
public sealed class HtmlLogicalDocument {
    private readonly Dictionary<HtmlLogicalNodeKind, int> _counts;
    private readonly List<string> _capabilities;

    internal HtmlLogicalDocument(HtmlLogicalNode root, IDictionary<HtmlLogicalNodeKind, int> counts, IEnumerable<string> capabilities) {
        Root = root ?? throw new ArgumentNullException(nameof(root));
        _counts = new Dictionary<HtmlLogicalNodeKind, int>(counts ?? throw new ArgumentNullException(nameof(counts)));
        _capabilities = (capabilities ?? throw new ArgumentNullException(nameof(capabilities)))
            .Where(capability => !string.IsNullOrWhiteSpace(capability))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(capability => capability, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>Root logical node.</summary>
    public HtmlLogicalNode Root { get; }

    /// <summary>Capability tags inferred from the logical document.</summary>
    public IReadOnlyList<string> Capabilities => _capabilities;

    /// <summary>Returns the number of nodes for a normalized kind.</summary>
    public int Count(HtmlLogicalNodeKind kind) {
        return _counts.TryGetValue(kind, out int count) ? count : 0;
    }

    /// <summary>Returns a snapshot of node counts by kind.</summary>
    public IReadOnlyDictionary<HtmlLogicalNodeKind, int> GetCounts() {
        return new Dictionary<HtmlLogicalNodeKind, int>(_counts);
    }
}
