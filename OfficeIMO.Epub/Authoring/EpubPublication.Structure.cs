using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Sets the XHTML body's EPUB 3 front/body/back-matter partition, retaining its other semantic tokens.
    /// This declares structure; reading order and navigation are controlled independently by the spine and SetNavigation.
    /// </summary>
    public void SetDocumentMatter(string manifestId, EpubDocumentMatter matter, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Document partition semantics require EPUB 3.");
        string token = matter switch {
            EpubDocumentMatter.FrontMatter => "frontmatter",
            EpubDocumentMatter.BodyMatter => "bodymatter",
            EpubDocumentMatter.BackMatter => "backmatter",
            _ => throw new ArgumentOutOfRangeException(nameof(matter))
        };
        XDocument content = EditableXhtml(manifestId);
        XElement body = content.Root!.Element(Html + "body") ?? throw new InvalidDataException("Content has no XHTML body.");
        body.SetAttributeValue(Ops + "type", string.Join(" ", Tokens((string?)body.Attribute(Ops + "type"))
            .Where(value => value != "frontmatter" && value != "bodymatter" && value != "backmatter")
            .Concat(new[] { token }).Distinct(StringComparer.Ordinal)));
        CommitContentEdits(new Dictionary<string, XDocument>(StringComparer.Ordinal) { [manifestId] = content }, cancellationToken);
    }

    /// <summary>
    /// Marks an existing empty XHTML span as a print-page boundary and appends its link to the EPUB 3 page list.
    /// Supply markers in the source edition's page order; labels may be Roman numerals or other publisher text.
    /// Content and navigation are committed atomically. Existing navigation and headings remain intact.
    /// This creates source-page references, not screen pagination or a print layout.
    /// </summary>
    public void AddPrintPageMarker(string manifestId, string markerId, string pageLabel, string pageListTitle = "Pages",
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Semantic print-page authoring requires EPUB 3.");
        RequireText(markerId, nameof(markerId)); RequireText(pageLabel, nameof(pageLabel)); RequireText(pageListTitle, nameof(pageListTitle));
        EpubManifestItem item = RequireManifestItem(manifestId);
        string path = RequireLocalPath(item);
        if (!Spine.Any(position => position.ManifestId == manifestId)) throw new InvalidDataException("Print-page markers must belong to a spine document.");
        XDocument content = EditableXhtml(manifestId);
        XElement marker = RequireContentElement(content, markerId);
        if (marker.Name != Html + "span" || marker.Elements().Any() || !string.IsNullOrWhiteSpace(marker.Value) || !marker.Ancestors(Html + "body").Any())
            throw new InvalidDataException("A print-page marker requires an empty XHTML span in the document body.");
        RequireCompatibleRole(marker, "doc-pagebreak");
        if (marker.Attribute("aria-labelledby") != null) throw new InvalidDataException("A print-page marker cannot retain a competing aria-labelledby name.");
        string navigationPath = NavigationPath();
        EpubManifestItem navigationItem = Manifest.Single(resource => resource.Reference.ContainerPath == navigationPath);
        XDocument navigation = navigationItem.Id == manifestId ? content : EditableXhtml(navigationItem.Id);
        XElement body = navigation.Root!.Element(Html + "body") ?? throw new InvalidDataException("Navigation has no XHTML body.");
        XElement[] pageLists = body.Descendants(Html + "nav").Where(element => HasToken((string?)element.Attribute(Ops + "type"), "page-list")).ToArray();
        if (pageLists.Length > 1) throw new InvalidDataException("Print-page authoring requires at most one page-list navigation section.");
        XElement nav;
        XElement list;
        if (pageLists.Length == 0) {
            list = new XElement(Html + "ol");
            nav = new XElement(Html + "nav", new XAttribute(Ops + "type", "page-list"),
                new XAttribute("role", "doc-pagelist"), new XElement(Html + "h1", pageListTitle), list);
            body.Add(nav);
        } else {
            nav = pageLists[0];
            RequireCompatibleRole(nav, "doc-pagelist");
            XElement[] lists = nav.Elements(Html + "ol").ToArray();
            if (lists.Length != 1) throw new InvalidDataException("Existing page-list navigation requires one ordered list.");
            list = lists[0];
        }
        string owner = HtmlContentLinkOwner(navigation, navigationPath);
        string? baseHref = navigation.Root!.Element(Html + "head")?.Elements(Html + "base")
            .Select(element => (string?)element.Attribute("href")).FirstOrDefault(value => value != null);
        foreach (XAttribute href in nav.Descendants(Html + "a").Attributes("href")) {
            EpubReference existing = EpubReference.Resolve(navigationPath, baseHref, href.Value);
            if (existing.ContainerPath == path && existing.Fragment == markerId)
                throw new InvalidDataException("A page-list entry already targets this marker.");
        }
        marker.SetAttributeValue("role", "doc-pagebreak");
        marker.SetAttributeValue("aria-label", pageLabel);
        AddSemanticToken(marker, "pagebreak");
        nav.SetAttributeValue("role", "doc-pagelist");
        list.Add(new XElement(Html + "li", new XElement(Html + "a",
            new XAttribute("href", RelativeHref(owner, path) + "#" + Uri.EscapeDataString(markerId)), pageLabel)));
        var documents = new Dictionary<string, XDocument>(StringComparer.Ordinal) { [manifestId] = content };
        documents[navigationItem.Id] = navigation;
        CommitContentEdits(documents, cancellationToken);
    }
}
