namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents a nested RTF group.
/// </summary>
public sealed class RtfGroup : RtfNode {
    internal RtfGroup(int position, IEnumerable<RtfNode> children)
        : base(position) {
        Children = new ReadOnlyCollection<RtfNode>((children ?? throw new ArgumentNullException(nameof(children))).ToList());
    }

    /// <summary>Child syntax nodes.</summary>
    public IReadOnlyList<RtfNode> Children { get; }

    /// <summary>
    /// Gets the first control word name in the group, if any.
    /// </summary>
    public string? Destination {
        get {
            foreach (RtfNode child in Children) {
                if (child is RtfControlWord control) {
                    return control.Name;
                }
            }

            return null;
        }
    }
}
