namespace OfficeIMO.OpenDocument;

/// <summary>A native ODF data-style definition with its original package scope.</summary>
public sealed class OdfDataStyle {
    internal OdfDataStyle(OdfDocument document, XElement element) { Document = document; Element = element; }
    internal OdfDocument Document { get; }
    internal XElement Element { get; }
    /// <summary>Name used by native data-style references.</summary>
    public string Name => (string?)Element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    /// <summary>Native definition kind, such as date-style, time-style, or number-style.</summary>
    public string ElementName => Element.Name.LocalName;
    /// <summary>Package part containing this definition.</summary>
    public string PartPath => Document.GetPartPath(Element);
    /// <summary>Whether the definition is part-local automatic styling.</summary>
    public bool IsAutomatic => Element.Parent?.Name == OdfNamespaces.Office + "automatic-styles";
    /// <summary>Returns an independent inspection copy of all native attributes and components.</summary>
    public XElement ToXml() => new(Element);
}
