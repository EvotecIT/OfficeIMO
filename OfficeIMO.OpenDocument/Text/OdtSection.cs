namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed named ODT section.</summary>
public sealed class OdtSection {
    private readonly OdtDocument _document;
    private readonly XElement _element;

    internal OdtSection(OdtDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>Section name.</summary>
    public string Name {
        get => (string?)_element.Attribute(OdfNamespaces.Text + "name") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Section name cannot be empty.", nameof(value));
            _element.SetAttributeValue(OdfNamespaces.Text + "name", value);
            Dirty();
        }
    }

    /// <summary>Paragraphs directly contained by the section.</summary>
    public IReadOnlyList<OdtParagraph> Paragraphs => _element.Elements()
        .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
        .Select(element => new OdtParagraph(_document, element)).ToList();

    /// <summary>Adds a paragraph to the section.</summary>
    public OdtParagraph AddParagraph(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdtTextCodec.Append(paragraph, text);
        _element.Add(paragraph);
        Dirty();
        return new OdtParagraph(_document, paragraph);
    }

    /// <summary>Adds a heading to the section.</summary>
    public OdtParagraph AddHeading(string text, int level = 1) {
        if (level < 1 || level > 10) throw new ArgumentOutOfRangeException(nameof(level));
        var heading = new XElement(OdfNamespaces.Text + "h", new XAttribute(OdfNamespaces.Text + "outline-level", level));
        OdtTextCodec.Append(heading, text);
        _element.Add(heading);
        Dirty();
        return new OdtParagraph(_document, heading);
    }

    private void Dirty() => _document.MarkPartDirty("content.xml");
}
