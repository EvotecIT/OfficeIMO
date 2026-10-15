namespace OfficeIMO.OpenDocument;

public sealed partial class OdgDocument {
    /// <summary>All existing drawing masters, including masters not referenced by a page. Reading does not create a master container.</summary>
    public IReadOnlyList<OdgMasterPage> MasterPages => GetExistingXml("styles.xml")?.Root?
        .Element(OdfNamespaces.Office + "master-styles")?.Elements(OdfNamespaces.Style + "master-page")
        .Select(element => new OdgMasterPage(this, element)).ToArray() ?? Array.Empty<OdgMasterPage>();
}

/// <summary>An existing native drawing master, with its own authored artwork and layer definitions.</summary>
public sealed class OdgMasterPage {
    private readonly OdgDocument _document;
    private readonly XElement _element;
    internal OdgMasterPage(OdgDocument document, XElement element) { _document = document; _element = element; }
    /// <summary>Native master name; inspection retains even an unnamed imported master.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    /// <summary>Native page layout reference.</summary>
    public string? PageLayoutName => (string?)_element.Attribute(OdfNamespaces.Style + "page-layout-name");
    /// <summary>Artwork in source paint order, including preserved drawing shapes.</summary>
    public OdgShapes Shapes => new OdgShapes(_document, _element);
    /// <summary>Layer definitions authored on this master.</summary>
    public OdgLayers Layers => new OdgLayers(_document, _element, "styles.xml");
}
