namespace OfficeIMO.Html.Dom;

/// <summary>Structural node categories retained independently of a parser implementation.</summary>
public enum HtmlNodeKind {
    /// <summary>An HTML document.</summary>
    Document,
    /// <summary>An element in the HTML, SVG, MathML or another namespace.</summary>
    Element,
    /// <summary>Decoded character data.</summary>
    Text,
    /// <summary>A source comment.</summary>
    Comment,
    /// <summary>A document type declaration.</summary>
    DocumentType,
    /// <summary>A detached group of nodes, including template contents.</summary>
    DocumentFragment
}
