namespace OfficeIMO.Html;

/// <summary>
/// Tracks shared HTML DOM traversal node and depth limits for converter packages.
/// </summary>
public sealed class HtmlDomLimitTracker {
    private readonly int? _maxHtmlNodes;
    private readonly int? _maxHtmlDepth;
    private int _nodes;

    /// <summary>
    /// Creates a DOM traversal limit tracker.
    /// </summary>
    /// <param name="maxHtmlNodes">Maximum number of traversed DOM nodes, or <c>null</c> for no node limit.</param>
    /// <param name="maxHtmlDepth">Maximum element nesting depth, or <c>null</c> for no depth limit.</param>
    public HtmlDomLimitTracker(int? maxHtmlNodes, int? maxHtmlDepth) {
        _maxHtmlNodes = maxHtmlNodes;
        _maxHtmlDepth = maxHtmlDepth;
    }

    /// <summary>
    /// Creates a tracker only when at least one DOM traversal limit is configured.
    /// </summary>
    /// <param name="maxHtmlNodes">Maximum number of traversed DOM nodes, or <c>null</c> for no node limit.</param>
    /// <param name="maxHtmlDepth">Maximum element nesting depth, or <c>null</c> for no depth limit.</param>
    /// <returns>A tracker when limits are configured; otherwise <c>null</c>.</returns>
    public static HtmlDomLimitTracker? Create(int? maxHtmlNodes, int? maxHtmlDepth) =>
        maxHtmlNodes.HasValue || maxHtmlDepth.HasValue
            ? new HtmlDomLimitTracker(maxHtmlNodes, maxHtmlDepth)
            : null;

    /// <summary>
    /// Records a traversed text or non-element node.
    /// </summary>
    public void RecordNode() {
        RecordNodeCount();
    }

    /// <summary>
    /// Records a traversed element start and validates its nesting depth.
    /// </summary>
    /// <param name="depth">One-based element nesting depth.</param>
    public void RecordElementStart(int depth) {
        RecordNodeCount();
        if (_maxHtmlDepth.HasValue && depth > _maxHtmlDepth.Value) {
            throw new HtmlDomLimitException(
                HtmlConversionDiagnosticCodes.HtmlDepthLimitExceeded,
                "HTML nesting depth exceeded the configured conversion limit.",
                "MaxHtmlDepth",
                depth,
                _maxHtmlDepth.Value);
        }
    }

    private void RecordNodeCount() {
        _nodes++;
        if (_maxHtmlNodes.HasValue && _nodes > _maxHtmlNodes.Value) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.NodeLimitExceeded,
                "HTML node count exceeded the configured conversion limit.",
                "MaxHtmlNodes",
                _nodes,
                _maxHtmlNodes.Value);
        }
    }
}
