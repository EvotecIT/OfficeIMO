namespace OfficeIMO.Html;

/// <summary>Semantic HTML list category.</summary>
public enum HtmlSemanticListKind {
    /// <summary>An unordered list.</summary>
    Unordered,
    /// <summary>An ordered list.</summary>
    Ordered,
    /// <summary>A definition list.</summary>
    Definition
}

/// <summary>Typed list state retained from an HTML list container.</summary>
public sealed class HtmlSemanticList {
    internal HtmlSemanticList(HtmlSemanticListKind kind, int? start, bool isReversed) {
        Kind = kind;
        Start = start;
        IsReversed = isReversed;
    }

    /// <summary>List category.</summary>
    public HtmlSemanticListKind Kind { get; }

    /// <summary>Effective first ordinal for an ordered list; otherwise null.</summary>
    public int? Start { get; }

    /// <summary>Whether ordered-list numbering descends.</summary>
    public bool IsReversed { get; }
}

/// <summary>Semantic HTML list-item category.</summary>
public enum HtmlSemanticListItemKind {
    /// <summary>A regular ordered or unordered list item.</summary>
    Item,
    /// <summary>A definition term.</summary>
    Term,
    /// <summary>A definition description.</summary>
    Description
}

/// <summary>Typed state retained from one HTML list item.</summary>
public sealed class HtmlSemanticListItem {
    internal HtmlSemanticListItem(HtmlSemanticListItemKind kind, int? ordinal, int? explicitOrdinal) {
        Kind = kind;
        Ordinal = ordinal;
        ExplicitOrdinal = explicitOrdinal;
    }

    /// <summary>List-item category.</summary>
    public HtmlSemanticListItemKind Kind { get; }

    /// <summary>Effective displayed ordinal for an ordered-list item; otherwise null.</summary>
    public int? Ordinal { get; }

    /// <summary>Explicit ordinal from a list-item value attribute; otherwise null.</summary>
    public int? ExplicitOrdinal { get; }
}
