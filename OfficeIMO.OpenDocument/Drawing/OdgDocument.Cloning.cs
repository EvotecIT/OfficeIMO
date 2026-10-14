namespace OfficeIMO.OpenDocument;

public sealed partial class OdgDocument {
    /// <summary>
    /// Appends a copy of a page in this drawing, remapping native identifiers and local references.
    /// Styles, immutable image entries and the master are shared. Clone the master separately for independent layout edits.
    /// </summary>
    /// <param name="sourceIndex">Zero-based source page position.</param>
    /// <param name="name">Unique destination name; null generates a name from the source page.</param>
    /// <exception cref="NotSupportedException">The page contains embedded editable objects, forms, animations or named text definitions.</exception>
    /// <exception cref="InvalidDataException">Identifiers or local attachment references are ambiguous or unresolved.</exception>
    public OdgPage ClonePage(int sourceIndex, string? name = null) {
        OdgPage source = Pages.ElementAtOrDefault(sourceIndex) ?? throw new ArgumentOutOfRangeException(nameof(sourceIndex));
        string destination = name ?? NextName(Pages.Select(page => page.Name), source.Name + "Copy");
        ValidateCloneName(destination, Pages.Select(page => page.Name), nameof(name));
        var context = new OdgCloneContext(GetXml("content.xml").Root!, GetXml("styles.xml").Root!);
        XElement clone = context.Clone(source.Element, source.Name, destination);
        DrawingBody.Add(clone); MarkPartDirty("content.xml");
        return new OdgPage(this, clone);
    }

    /// <summary>
    /// Clones a master and its page layout in this drawing and returns its unique native name.
    /// Artwork and layers are copied; styles and immutable image entries remain shared.
    /// Assign the result to a page's MasterPageName to use the independent layout and layer set.
    /// </summary>
    /// <param name="sourceName">Existing, unambiguous master name.</param>
    /// <param name="name">Unique destination name; null generates a name from the source master.</param>
    /// <exception cref="NotSupportedException">The master contains content outside the page-cloning profile.</exception>
    /// <exception cref="InvalidDataException">The source master or page layout is missing or ambiguous.</exception>
    public string CloneMasterPage(string sourceName, string? name = null) {
        if (string.IsNullOrWhiteSpace(sourceName)) throw new ArgumentException("Master name cannot be empty.", nameof(sourceName));
        XElement styles = GetXml("styles.xml").Root!;
        XElement masters = styles.Element(OdfNamespaces.Office + "master-styles")
            ?? throw new InvalidDataException("Drawing has no master styles.");
        XElement source = UniqueCloneSource(masters.Elements(OdfNamespaces.Style + "master-page"), sourceName, "master");
        string destination = name ?? NextName(masters.Elements().Select(StyleName), sourceName + "Copy");
        OdfStyleRepository.ValidateStyleName(destination);
        ValidateCloneName(destination, masters.Elements().Select(StyleName), nameof(name));
        string layoutName = (string?)source.Attribute(OdfNamespaces.Style + "page-layout-name")
            ?? throw new InvalidDataException("Drawing master has no page layout.");
        XElement automatic = styles.Element(OdfNamespaces.Office + "automatic-styles")
            ?? throw new InvalidDataException("Drawing has no automatic styles.");
        XElement layout = UniqueCloneSource(automatic.Elements(OdfNamespaces.Style + "page-layout"), layoutName, "page layout");
        if (layout.Element(OdfNamespaces.Style + "page-layout-properties") == null)
            throw new InvalidDataException("Drawing master page layout has no properties.");
        string destinationLayout = NextName(automatic.Elements().Select(StyleName), layoutName + "Copy");
        var context = new OdgCloneContext(GetXml("content.xml").Root!, styles);
        XElement layoutClone = context.Clone(layout);
        XElement masterClone = context.Clone(source);
        layoutClone.SetAttributeValue(OdfNamespaces.Style + "name", destinationLayout);
        masterClone.SetAttributeValue(OdfNamespaces.Style + "name", destination);
        masterClone.SetAttributeValue(OdfNamespaces.Style + "page-layout-name", destinationLayout);
        if ((string?)masterClone.Attribute(OdfNamespaces.Style + "next-style-name") == sourceName)
            masterClone.SetAttributeValue(OdfNamespaces.Style + "next-style-name", destination);
        automatic.AddFirst(layoutClone); masters.Add(masterClone); MarkPartDirty("styles.xml");
        return destination;
    }

    private static string StyleName(XElement element) => (string?)element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    private static XElement UniqueCloneSource(IEnumerable<XElement> elements, string name, string kind) {
        XElement[] matches = elements.Where(element => StyleName(element) == name).Take(2).ToArray();
        return matches.Length == 1 ? matches[0] : throw new InvalidDataException("Drawing " + kind + " name is missing or ambiguous: " + name);
    }
    private static void ValidateCloneName(string name, IEnumerable<string> existing, string parameter) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Clone name cannot be empty.", parameter);
        if (existing.Contains(name, StringComparer.Ordinal)) throw new ArgumentException("Clone name must be unique.", parameter);
    }
}
