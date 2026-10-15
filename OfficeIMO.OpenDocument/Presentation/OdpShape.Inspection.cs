namespace OfficeIMO.OpenDocument;

public abstract partial class OdpShape {
    /// <summary>Paragraphs in this shape's own text story, including list items. Group children and table cells have separate stories.</summary>
    public IReadOnlyList<OdpParagraph> TextParagraphs => this is OdpTextBox textBox ? textBox.Paragraphs : OdfTextTraversal.Paragraphs(
        Element.Element(OdfNamespaces.Draw + "text-box") ?? Element)
        .Select(element => new OdpParagraph(Presentation, element)).ToArray();
}
