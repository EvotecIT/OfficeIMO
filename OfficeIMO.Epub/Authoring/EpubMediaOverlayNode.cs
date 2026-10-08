namespace OfficeIMO.Epub;

/// <summary>A text-associated narration cue or nested sequence.</summary>
public abstract class EpubMediaOverlayNode {
    internal EpubMediaOverlayNode(string elementId) => ElementId = elementId ?? throw new ArgumentNullException(nameof(elementId));
    /// <summary>Existing unqualified id of an XHTML body element or a supported SVG element.</summary>
    public string ElementId { get; }
    /// <summary>Optional structural meaning offered to reading systems for skipping or escaping.</summary>
    public EpubMediaOverlaySemantic? Semantic { get; set; }
}

/// <summary>Structural semantics supported by authored narration. Reader controls are implementation-dependent.</summary>
public enum EpubMediaOverlaySemantic {
    /// <summary>A footnote.</summary>
    Footnote,
    /// <summary>An endnote.</summary>
    Endnote,
    /// <summary>A print page boundary.</summary>
    PageBreak,
    /// <summary>A table.</summary>
    Table,
    /// <summary>A table row.</summary>
    TableRow,
    /// <summary>A table cell.</summary>
    TableCell,
    /// <summary>A list.</summary>
    List,
    /// <summary>A list item.</summary>
    ListItem,
    /// <summary>A figure.</summary>
    Figure,
    /// <summary>Secondary content.</summary>
    Aside
}
