namespace OfficeIMO.OpenDocument;

public sealed partial class OdtDocument {
    /// <summary>Existing header/footer stories across all authored masters, without creating a default layout or missing stories.</summary>
    public IReadOnlyList<OdtHeaderFooter> HeaderFooters => GetExistingXml("styles.xml")?.Root?
        .Element(OdfNamespaces.Office + "master-styles")?.Elements(OdfNamespaces.Style + "master-page")
        .SelectMany(master => master.Elements().Where(element => element.Name.Namespace == OdfNamespaces.Style &&
            element.Name.LocalName is "header" or "footer" or "header-first" or "footer-first" or "header-left" or "footer-left"))
        .Select(element => new OdtHeaderFooter(this, element)).ToArray() ?? Array.Empty<OdtHeaderFooter>();
}
