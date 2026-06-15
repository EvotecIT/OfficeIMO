namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Base class for syntax-tree nodes.
/// </summary>
public abstract class RtfNode {
    private protected RtfNode(int position) {
        Position = position;
    }

    /// <summary>Zero-based source position.</summary>
    public int Position { get; }
}
