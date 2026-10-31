namespace OfficeIMO.Word;

/// <summary>An editable node attached to a Word SmartArt diagram.</summary>
public sealed class WordSmartArtNode {
    internal WordSmartArtNode(WordSmartArt smartArt, string id) {
        SmartArt = smartArt;
        Id = id;
    }

    /// <summary>Gets the diagram containing this node.</summary>
    public WordSmartArt SmartArt { get; }

    /// <summary>Gets the native node identity, which remains stable when nodes are reordered.</summary>
    public string Id { get; }

    /// <summary>Gets or replaces the node text while retaining its paragraph end properties.</summary>
    /// <exception cref="InvalidOperationException">The node has been removed from its diagram.</exception>
    public string Text {
        get => SmartArt.ReadNodeText(Id);
        set => SmartArt.WriteNodeText(Id, value);
    }
}
