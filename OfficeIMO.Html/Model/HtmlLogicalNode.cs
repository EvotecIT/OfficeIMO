namespace OfficeIMO.Html;

/// <summary>
/// Normalized node in an OfficeIMO HTML logical document.
/// </summary>
public sealed class HtmlLogicalNode {
    private readonly List<HtmlLogicalNode> _children = new List<HtmlLogicalNode>();
    private readonly Dictionary<string, string> _attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _capabilities = new List<string>();

    internal HtmlLogicalNode(HtmlLogicalNodeKind kind, string name, string text) {
        Kind = kind;
        Name = name ?? string.Empty;
        Text = text ?? string.Empty;
    }

    /// <summary>Normalized node kind.</summary>
    public HtmlLogicalNodeKind Kind { get; }

    /// <summary>Original element name, or <c>#text</c> for text nodes.</summary>
    public string Name { get; }

    /// <summary>Trimmed node text for text and small semantic nodes.</summary>
    public string Text { get; }

    /// <summary>Attributes captured from the source element.</summary>
    public IReadOnlyDictionary<string, string> Attributes => _attributes;

    /// <summary>Capability tags inferred from this node.</summary>
    public IReadOnlyList<string> Capabilities => _capabilities;

    /// <summary>Child logical nodes.</summary>
    public IReadOnlyList<HtmlLogicalNode> Children => _children;

    internal void AddAttribute(string name, string value) {
        if (!string.IsNullOrWhiteSpace(name)) {
            _attributes[name] = value ?? string.Empty;
        }
    }

    internal void AddCapability(string capability) {
        if (!string.IsNullOrWhiteSpace(capability) && !_capabilities.Contains(capability)) {
            _capabilities.Add(capability);
        }
    }

    internal void AddChild(HtmlLogicalNode child) {
        if (child != null) {
            _children.Add(child);
        }
    }
}
