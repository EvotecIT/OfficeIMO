namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed basic ODP animation effect.</summary>
public sealed class OdpAnimation {
    private readonly OdpPresentation _presentation;
    private readonly XElement _element;
    internal OdpAnimation(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>Target shape XML identifier.</summary>
    public string TargetElement => (string?)_element.Attribute(OdfNamespaces.Smil + "targetElement") ?? string.Empty;
    /// <summary>Animated attribute name.</summary>
    public string AttributeName => (string?)_element.Attribute(OdfNamespaces.Smil + "attributeName") ?? string.Empty;
    /// <summary>Starting value.</summary>
    public string? From => (string?)_element.Attribute(OdfNamespaces.Smil + "from");
    /// <summary>Ending value.</summary>
    public string? To => (string?)_element.Attribute(OdfNamespaces.Smil + "to");
    /// <summary>Effect duration.</summary>
    public TimeSpan Duration {
        get {
            string? lexical = (string?)_element.Attribute(OdfNamespaces.Smil + "dur");
            try { return lexical == null ? TimeSpan.Zero : XmlConvert.ToTimeSpan(lexical); }
            catch (FormatException) { return TimeSpan.Zero; }
        }
    }
    /// <summary>Removes this effect.</summary>
    public void Remove() {
        XElement? parent = _element.Parent;
        _element.Remove();
        while (parent != null && parent.Name.Namespace == OdfNamespaces.Anim && !parent.Elements().Any()) {
            XElement? next = parent.Parent; parent.Remove(); parent = next;
        }
        _presentation.MarkPartDirty("content.xml");
    }
}
