namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    /// <summary>Native paragraphs and headings in reading order, including list items. Excludes other text stories.</summary>
    public IReadOnlyList<OdfTextParagraph> Paragraphs {
        get { RequireTextEditing(); return OdfTextTraversal.Paragraphs(TextRoot).Select(e => new OdfTextParagraph(Document, e, Element)).ToList(); }
    }
    /// <summary>Direct native lists in this shape's text container.</summary>
    public IReadOnlyList<OdfTextList> Lists {
        get { RequireTextEditing(); return TextRoot.Elements(OdfNamespaces.Text + "list").Select(e => new OdfTextList(Document, e, Element)).ToList(); }
    }
    /// <summary>Appends a native paragraph while retaining existing text, lists and shape metadata.</summary>
    public OdfTextParagraph AddParagraph(string? text = null) {
        RequireTextEditing();
        var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, text);
        OdfDrawTextInsertion.Append(TextRoot, paragraph); Dirty(); return new OdfTextParagraph(Document, paragraph, Element);
    }
    /// <summary>Appends a native ordered or unordered list with its own automatic list style.</summary>
    public OdfTextList AddList(bool ordered = false) {
        RequireTextEditing();
        string style = OdfListStyleStore.Create(Document, ordered, PartPath);
        var list = new XElement(OdfNamespaces.Text + "list", new XAttribute(OdfNamespaces.Text + "style-name", style));
        OdfDrawTextInsertion.Append(TextRoot, list); Dirty(); return new OdfTextList(Document, list, Element);
    }
    private void RequireTextEditing() {
        if (IsGroup || !new[] { "rect", "ellipse", "circle", "frame", "custom-shape", "path", "polygon", "polyline", "line", "connector" }.Contains(ElementName))
            throw new NotSupportedException("Text editing is not supported for this element.");
        if (ElementName == "frame" && Element.Element(OdfNamespaces.Draw + "text-box") == null && !IsImage)
            throw new NotSupportedException("This frame has no text box or image caption.");
    }
}
