using System.Globalization;

namespace OfficeIMO.OpenDocument;

/// <summary>A spreadsheet cell annotation (the OpenDocument equivalent of an Excel note/comment).</summary>
public sealed class OdsAnnotation {
    private readonly OdsDocument _document;
    private readonly XElement _element;

    internal OdsAnnotation(OdsDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>Stable annotation name when one was authored.</summary>
    public string? Name {
        get => (string?)_element.Attribute(OdfNamespaces.Office + "name");
        set => SetAttribute(OdfNamespaces.Office + "name", value);
    }

    /// <summary>Annotation author.</summary>
    public string? Creator {
        get => (string?)_element.Element(OdfNamespaces.Dc + "creator");
        set => SetElement(OdfNamespaces.Dc + "creator", value);
    }

    /// <summary>Annotation timestamp, when represented by a valid round-trip ODF date.</summary>
    public DateTimeOffset? Date {
        get {
            string? value = (string?)_element.Element(OdfNamespaces.Dc + "date");
            return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTimeOffset parsed)
                ? parsed
                : (DateTimeOffset?)null;
        }
        set => SetElement(OdfNamespaces.Dc + "date", value?.ToString("o", CultureInfo.InvariantCulture));
    }

    /// <summary>Plain annotation body with ODF spaces, tabs, and line breaks decoded.</summary>
    public string Text {
        get => string.Join("\n", ReadTextBlocks(_element));
        set {
            if (value == null) throw new ArgumentNullException(nameof(value));
            _element.Elements().Where(IsAnnotationTextBlock).Remove();
            var paragraph = new XElement(OdfNamespaces.Text + "p");
            OdfTextCodec.Append(paragraph, value);
            _element.Add(paragraph);
            Dirty();
        }
    }

    /// <summary>Removes this annotation from its cell.</summary>
    public void Remove() {
        if (_element.Parent == null) return;
        _element.Remove();
        Dirty();
    }

    private void SetAttribute(XName name, string? value) {
        _element.SetAttributeValue(name, string.IsNullOrWhiteSpace(value) ? null : value);
        Dirty();
    }

    private void SetElement(XName name, string? value) {
        XElement? element = _element.Element(name);
        if (string.IsNullOrWhiteSpace(value)) {
            element?.Remove();
        } else if (element == null) {
            XElement? firstParagraph = _element.Elements().FirstOrDefault(IsAnnotationTextBlock);
            var added = new XElement(name, value);
            XElement? date = _element.Element(OdfNamespaces.Dc + "date");
            if (name == OdfNamespaces.Dc + "creator" && date != null) date.AddBeforeSelf(added);
            else if (firstParagraph == null) _element.Add(added);
            else firstParagraph.AddBeforeSelf(added);
        } else {
            element.Value = value;
        }
        EnsureCreatorPrecedesDate();
        Dirty();
    }

    private void EnsureCreatorPrecedesDate() {
        XElement? creator = _element.Element(OdfNamespaces.Dc + "creator");
        XElement? date = _element.Element(OdfNamespaces.Dc + "date");
        if (creator == null || date == null
            || creator.NodesAfterSelf().Contains(date)) return;
        creator.Remove();
        date.AddBeforeSelf(creator);
    }

    private static IEnumerable<string> ReadTextBlocks(XElement container) {
        foreach (XElement child in container.Elements()) {
            if (child.Name == OdfNamespaces.Text + "p" || child.Name == OdfNamespaces.Text + "h") {
                yield return OdfTextCodec.Read(child);
            } else if (child.Name == OdfNamespaces.Text + "list") {
                foreach (string block in ReadListBlocks(child)) yield return block;
            }
        }
    }

    private static IEnumerable<string> ReadListBlocks(XElement list) {
        foreach (XElement child in list.Elements()) {
            if (child.Name != OdfNamespaces.Text + "list-item" &&
                child.Name != OdfNamespaces.Text + "list-header") continue;
            foreach (string block in ReadTextBlocks(child)) yield return block;
        }
    }

    private static bool IsAnnotationTextBlock(XElement element) =>
        element.Name == OdfNamespaces.Text + "p" ||
        element.Name == OdfNamespaces.Text + "h" ||
        element.Name == OdfNamespaces.Text + "list";

    private void Dirty() => _document.MarkPartDirty("content.xml");
}
