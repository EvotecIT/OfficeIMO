namespace OfficeIMO.OpenDocument;

public sealed partial class OdgDocument {
    /// <summary>
    /// Appends a page from another drawing, with independent master/layout, reachable styles and embedded resources.
    /// Native identifiers are remapped. Unsupported content or unresolved dependencies fail before changing either document.
    /// </summary>
    /// <param name="source">A different drawing, loaded from ODG or FODG.</param>
    /// <param name="sourceIndex">Zero-based source page position.</param>
    /// <param name="name">Unique destination name; null generates a name from the source page.</param>
    /// <exception cref="NotSupportedException">Content or default-style inheritance cannot be transferred without changing its meaning.</exception>
    /// <exception cref="InvalidDataException">A required master, layout, style, resource or reference is missing or ambiguous.</exception>
    public OdgPage ImportPage(OdgDocument source, int sourceIndex, string? name = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (ReferenceEquals(this, source)) throw new ArgumentException("Use ClonePage for copies within this drawing.", nameof(source));
        if (source.Version > Version) throw new NotSupportedException("Drawing import requires a destination using the source ODF version or newer.");
        OdgPage page = source.Pages.ElementAtOrDefault(sourceIndex) ?? throw new ArgumentOutOfRangeException(nameof(sourceIndex));
        string destinationName = name ?? NextName(Pages.Select(item => item.Name), page.Name + "Copy");
        ValidateCloneName(destinationName, Pages.Select(item => item.Name), nameof(name));
        XElement master = page.Master ?? throw new InvalidDataException("Imported drawing page has no master.");
        string layoutName = (string?)master.Attribute(OdfNamespaces.Style + "page-layout-name")
            ?? throw new InvalidDataException("Imported drawing master has no page layout.");
        XElement sourceStyles = source.GetXml("styles.xml").Root!;
        XElement layout = UniqueCloneSource(sourceStyles.Element(OdfNamespaces.Office + "automatic-styles")?
            .Elements(OdfNamespaces.Style + "page-layout") ?? Enumerable.Empty<XElement>(), layoutName, "page layout");
        if (layout.Element(OdfNamespaces.Style + "page-layout-properties") == null)
            throw new InvalidDataException("Imported drawing page layout has no properties.");
        var cloneContext = new OdgCloneContext(GetXml("content.xml").Root!, GetXml("styles.xml").Root!);
        XElement[] forest = cloneContext.CloneForest(new[] { page.Element, master, layout }, page.Name, destinationName);
        XElement pageCopy = forest[0], masterCopy = forest[1], layoutCopy = forest[2];
        pageCopy.SetAttributeValue(OdfNamespaces.Draw + "name", destinationName);
        var resources = new OdfResourceImportPlan(source.Package, Package);
        var styles = new OdfStyleImportPlan(source, this, resources);
        string masterName = styles.ReserveName(), destinationLayout = styles.ReserveName();
        masterCopy.SetAttributeValue(OdfNamespaces.Style + "name", masterName);
        masterCopy.SetAttributeValue(OdfNamespaces.Style + "page-layout-name", destinationLayout);
        layoutCopy.SetAttributeValue(OdfNamespaces.Style + "name", destinationLayout);
        pageCopy.SetAttributeValue(OdfNamespaces.Draw + "master-page-name", masterName);
        styles.SetMasterMapping(page.MasterPageName, masterName);
        styles.RewriteContent(page.Element, pageCopy, "content.xml"); styles.RewriteContent(master, masterCopy, "styles.xml"); styles.Rewrite(layoutCopy, "styles.xml");
        SnapshotLayers(page, pageCopy);
        resources.Rewrite(pageCopy); resources.Rewrite(masterCopy); resources.Rewrite(layoutCopy);
        // All fallible dependency resolution operates on detached XML and staged bytes above.
        resources.Apply(); styles.Apply();
        XElement destinationStyles = GetXml("styles.xml").Root!;
        // Keep added layouts ahead of styles when repeated imports append their dependency styles.
        EnsureContainer(destinationStyles, OdfNamespaces.Office + "automatic-styles").AddFirst(new XElement(layoutCopy));
        EnsureContainer(destinationStyles, OdfNamespaces.Office + "master-styles").Add(new XElement(masterCopy));
        DrawingBody.Add(new XElement(pageCopy)); MarkPartDirty("styles.xml"); MarkPartDirty("content.xml");
        return Pages.Last();
    }

    private static void SnapshotLayers(OdgPage source, XElement copy) {
        OdgLayers effective = source.EffectiveLayers;
        if (effective.GroupBy(layer => layer.Name, StringComparer.Ordinal).Any(group => group.Count() > 1))
            throw new InvalidDataException("Imported layer names are ambiguous.");
        XElement? set = copy.Element(OdfNamespaces.Draw + "layer-set");
        if (set == null) {
            set = new XElement(OdfNamespaces.Draw + "layer-set");
            XElement? description = copy.Elements().LastOrDefault(element => element.Name == OdfNamespaces.Svg + "title" || element.Name == OdfNamespaces.Svg + "desc");
            if (description == null) copy.AddFirst(set); else description.AddAfterSelf(set);
            foreach (OdgLayer layer in effective) set.Add(new XElement(OdfNamespaces.Draw + "layer", new XAttribute(OdfNamespaces.Draw + "name", layer.Name)));
        }
        foreach (XElement layer in set.Elements(OdfNamespaces.Draw + "layer")) {
            OdgLayer original = effective.Find((string?)layer.Attribute(OdfNamespaces.Draw + "name") ?? "")
                ?? throw new InvalidDataException("Imported layer declaration is ambiguous or missing.");
            layer.SetAttributeValue(OdfNamespaces.Draw + "display", original.Display.ToString().ToLowerInvariant());
            layer.SetAttributeValue(OdfNamespaces.Draw + "protected", original.IsProtected ? "true" : "false");
        }
    }
}
