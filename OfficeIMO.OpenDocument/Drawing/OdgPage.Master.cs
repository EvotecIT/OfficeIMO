namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    /// <summary>
    /// Referenced master name, or an empty string when absent in imported XML.
    /// Assigning an existing master shares its layout and inherited layers without changing page shapes or page layers.
    /// </summary>
    /// <exception cref="ArgumentException">The name is empty or does not identify a master in this document.</exception>
    /// <exception cref="InvalidDataException">The master name is ambiguous or its page layout cannot be resolved.</exception>
    public string MasterPageName {
        get => (string?)Element.Attribute(OdfNamespaces.Draw + "master-page-name") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Master page name cannot be empty.", nameof(value));
            XElement master = FindMaster(value) ?? throw new ArgumentException("Master page does not exist in this drawing.", nameof(value));
            ResolveLayoutProperties(master);
            Element.SetAttributeValue(OdfNamespaces.Draw + "master-page-name", value);
            _document.MarkPartDirty("content.xml");
        }
    }

    /// <summary>
    /// Layers declared on the referenced master. Adding a master layer set overrides document layers for pages using this master.
    /// An explicit page layer set continues to take precedence. Edits affect every page inheriting this master.
    /// </summary>
    /// <exception cref="InvalidDataException">The page does not reference a resolvable master.</exception>
    public OdgLayers MasterLayers => new OdgLayers(_document,
        Master ?? throw new InvalidDataException("Drawing page has no resolvable master."), "styles.xml");

    /// <summary>
    /// Artwork on the referenced master, in paint order. Edits affect every page using this master.
    /// Clone the master and assign its name to this page before making independent artwork edits.
    /// Unknown drawing elements remain visible and preserved outside the editable profile.
    /// </summary>
    /// <exception cref="InvalidDataException">The page does not reference a resolvable master.</exception>
    public OdgShapes MasterShapes => new OdgShapes(_document,
        Master ?? throw new InvalidDataException("Drawing page has no resolvable master."));

    private XElement? FindMaster(string name) {
        XElement[] matches = _document.GetXml("styles.xml").Root?.Element(OdfNamespaces.Office + "master-styles")?
            .Elements(OdfNamespaces.Style + "master-page")
            .Where(master => (string?)master.Attribute(OdfNamespaces.Style + "name") == name).Take(2).ToArray() ?? Array.Empty<XElement>();
        if (matches.Length > 1) throw new InvalidDataException("Drawing master page name is ambiguous.");
        return matches.FirstOrDefault();
    }

    private XElement ResolveLayoutProperties(XElement? master) {
        string? name = (string?)master?.Attribute(OdfNamespaces.Style + "page-layout-name");
        XElement[] matches = _document.GetXml("styles.xml").Root?.Element(OdfNamespaces.Office + "automatic-styles")?
            .Elements(OdfNamespaces.Style + "page-layout")
            .Where(layout => name != null && (string?)layout.Attribute(OdfNamespaces.Style + "name") == name).Take(2).ToArray() ?? Array.Empty<XElement>();
        if (matches.Length != 1 || matches[0].Element(OdfNamespaces.Style + "page-layout-properties") is not XElement properties)
            throw new InvalidDataException("Drawing page has no unambiguous page layout.");
        return properties;
    }
}
