namespace OfficeIMO.OpenDocument;

/// <summary>A drawing page. Page dimensions resolve through its referenced master and layout.</summary>
public sealed partial class OdgPage {
    private readonly OdgDocument _document;
    internal XElement Element { get; }
    internal OdgPage(OdgDocument document, XElement element) { _document = document; Element = element; }
    /// <summary>Unique page name.</summary>
    public string Name {
        get => (string?)Element.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Page name cannot be empty.", nameof(value));
            if (_document.Pages.Any(page => !ReferenceEquals(page.Element, Element) && page.Name == value)) throw new ArgumentException("Page names must be unique.", nameof(value));
            Element.SetAttributeValue(OdfNamespaces.Draw + "name", value); _document.MarkPartDirty("content.xml");
        }
    }
    /// <summary>Shapes in paint order, including preserved shapes outside the editable profile.</summary>
    public OdgShapes Shapes => new OdgShapes(_document, Element);
    /// <summary>Layers declared on this page. An explicit page set overrides master and document layers.</summary>
    public OdgLayers Layers => new OdgLayers(_document, Element, "content.xml");
    /// <summary>Layer definitions after page, master, and document inheritance.</summary>
    public OdgLayers EffectiveLayers {
        get {
            if (Layers.IsDeclared) return Layers;
            if (Master is XElement master) {
                var layers = new OdgLayers(_document, master, "styles.xml");
                if (layers.IsDeclared) return layers;
            }
            return _document.Layers;
        }
    }
    /// <summary>Page width. Changing a shared layout affects all pages referencing it.</summary>
    public OdfLength Width { get => ReadDimension("page-width"); set => SetDimension("page-width", value); }
    /// <summary>Page height. Changing a shared layout affects all pages referencing it.</summary>
    public OdfLength Height { get => ReadDimension("page-height"); set => SetDimension("page-height", value); }
    internal XElement? Master => FindMaster(MasterPageName);
    private XElement LayoutProperties => ResolveLayoutProperties(Master);
    private OdfLength ReadDimension(string name) => OdfLength.Parse((string?)LayoutProperties.Attribute(OdfNamespaces.Fo + name)
        ?? throw new InvalidDataException("Page layout is missing " + name + "."));
    private void SetDimension(string name, OdfLength value) {
        OdgDocument.ValidateDimension(value); LayoutProperties.SetAttributeValue(OdfNamespaces.Fo + name, value); _document.MarkPartDirty("styles.xml");
    }
}
