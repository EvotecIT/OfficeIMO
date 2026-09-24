namespace OfficeIMO.OpenDocument;

/// <summary>Native ODT note category.</summary>
public enum OdtNoteKind {
    /// <summary>A note at the foot of the page.</summary>
    Footnote,
    /// <summary>A note at the end of the document or section.</summary>
    Endnote
}

/// <summary>An XML-backed ODT footnote or endnote at its inline reference position.</summary>
public sealed class OdtNote {
    private readonly OdtDocument _document;
    private readonly XElement _element;

    internal OdtNote(OdtDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>The native note identifier.</summary>
    public string? Id => (string?)_element.Attribute(OdfNamespaces.Text + "id");

    /// <summary>The note kind, or null when the native class is unknown.</summary>
    public OdtNoteKind? Kind => (string?)_element.Attribute(OdfNamespaces.Text + "note-class") switch {
        "footnote" => OdtNoteKind.Footnote,
        "endnote" => OdtNoteKind.Endnote,
        _ => null
    };

    /// <summary>The displayed citation stored in the native note.</summary>
    public string Citation => _element.Element(OdfNamespaces.Text + "note-citation")?.Value ?? string.Empty;

    /// <summary>Direct paragraphs in the note body, in source order.</summary>
    public IReadOnlyList<OdtParagraph> Paragraphs => Body?.Elements(OdfNamespaces.Text + "p")
        .Select(paragraph => new OdtParagraph(_document, paragraph)).ToList() ?? new List<OdtParagraph>();

    /// <summary>Appends a paragraph to this note's body.</summary>
    public OdtParagraph AddParagraph(string? text = null) {
        XElement body = Body ?? throw new InvalidDataException("The ODT note has no note body.");
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        body.Add(paragraph);
        _document.MarkPartDirty("content.xml");
        return new OdtParagraph(_document, paragraph);
    }

    /// <summary>Whether the body contains only direct paragraphs that the Word adapter can project.</summary>
    public bool HasOnlyParagraphs => Body != null && Body.Elements().All(child => child.Name == OdfNamespaces.Text + "p");

    internal static OdtNote Create(OdtDocument document, OdtNoteKind kind, string id, string citation, string? text) {
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        var element = new XElement(OdfNamespaces.Text + "note",
            new XAttribute(OdfNamespaces.Text + "id", id),
            new XAttribute(OdfNamespaces.Text + "note-class", kind == OdtNoteKind.Footnote ? "footnote" : "endnote"),
            new XElement(OdfNamespaces.Text + "note-citation", citation),
            new XElement(OdfNamespaces.Text + "note-body", paragraph));
        return new OdtNote(document, element);
    }

    internal XElement Element => _element;
    private XElement? Body => _element.Element(OdfNamespaces.Text + "note-body");
}
