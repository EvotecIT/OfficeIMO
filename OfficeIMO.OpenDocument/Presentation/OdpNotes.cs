namespace OfficeIMO.OpenDocument;

/// <summary>XML-backed speaker notes for one slide.</summary>
public sealed class OdpNotes {
    private readonly OdpPresentation _presentation; private readonly XElement _element;
    internal OdpNotes(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>Speaker-note paragraphs.</summary>
    public IReadOnlyList<OdpParagraph> Paragraphs => _element.Descendants(OdfNamespaces.Text + "p").Select(element => new OdpParagraph(_presentation, element)).ToList();
    /// <summary>Adds a speaker-note paragraph.</summary>
    public OdpParagraph AddParagraph(string? text = null) {
        XElement? textBox = _element.Descendants(OdfNamespaces.Draw + "text-box").FirstOrDefault();
        if (textBox == null) {
            textBox = new XElement(OdfNamespaces.Draw + "text-box");
            _element.Add(new XElement(OdfNamespaces.Draw + "frame",
                new XAttribute(OdfNamespaces.Draw + "name", "Notes"), new XAttribute(OdfNamespaces.Presentation + "class", "notes"), textBox));
        }
        var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, text); textBox.Add(paragraph);
        _presentation.MarkPartDirty("content.xml"); return new OdpParagraph(_presentation, paragraph);
    }
}
