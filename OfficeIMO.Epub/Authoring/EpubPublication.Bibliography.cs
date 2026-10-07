using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Appends a formatted XHTML reference to an existing ol/ul directly inside a body section.
    /// Entries retain caller order. Citation formatting belongs to the bibliography renderer or publisher;
    /// this operation adds EPUB bibliography structure without sorting, numbering citations, or fetching sources.
    /// </summary>
    public void AddBibliographyEntry(string manifestId, string listId, string entryId, string entryXhtml,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Semantic bibliography authoring requires EPUB 3.");
        RequireText(listId, nameof(listId)); RequireText(entryXhtml, nameof(entryXhtml));
        XmlConvert.VerifyNCName(entryId);
        XDocument content = EditableXhtml(manifestId);
        XElement list = RequireContentElement(content, listId);
        XElement section = RequireBibliographyList(list);
        RequireCompatibleRole(section, "doc-bibliography");
        list.Add(new XElement(Html + "li", new XAttribute("id", entryId), ParseAuthoringFragment(entryXhtml)));
        AddSemanticToken(section, "bibliography");
        section.SetAttributeValue("role", "doc-bibliography");
        CommitContentEdits(new Dictionary<string, XDocument>(StringComparer.Ordinal) { [manifestId] = content }, cancellationToken);
    }

    /// <summary>Links an existing labelled body anchor to a bibliography entry, with an optional localized return link.</summary>
    public void LinkBibliographyEntry(string sourceManifestId, string referenceId, string bibliographyManifestId, string entryId,
        string? backlinkText = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Semantic bibliography authoring requires EPUB 3.");
        RequireText(referenceId, nameof(referenceId)); RequireText(entryId, nameof(entryId));
        if (backlinkText != null) RequireText(backlinkText, nameof(backlinkText));
        XDocument source = EditableXhtml(sourceManifestId);
        XDocument bibliography = sourceManifestId == bibliographyManifestId ? source : EditableXhtml(bibliographyManifestId);
        XElement marker = RequireContentElement(source, referenceId);
        if (marker.Name != Html + "a" || marker.Attribute("href") != null || marker.Ancestors(Html + "a").Any() ||
            marker.Descendants(Html + "a").Any() || !marker.Ancestors(Html + "body").Any())
            throw new InvalidDataException("A citation reference requires a body anchor without an href or nested anchor.");
        if (string.IsNullOrWhiteSpace(marker.Value) && string.IsNullOrWhiteSpace((string?)marker.Attribute("aria-label")))
            throw new InvalidDataException("A citation reference needs visible text or an aria-label.");
        RequireCompatibleRole(marker, "doc-biblioref");
        XElement entry = RequireContentElement(bibliography, entryId);
        if (entry.Name != Html + "li" || entry.Parent == null) throw new InvalidDataException("A bibliography target requires a list item.");
        XElement section = RequireBibliographyList(entry.Parent);
        if (!HasToken((string?)section.Attribute("role"), "doc-bibliography") || !HasToken((string?)section.Attribute(Ops + "type"), "bibliography"))
            throw new InvalidDataException("The target must belong to a semantic bibliography section.");
        string sourcePath = RequireLocalPath(RequireManifestItem(sourceManifestId));
        string bibliographyPath = RequireLocalPath(RequireManifestItem(bibliographyManifestId));
        marker.SetAttributeValue("href", RelativeHref(HtmlContentLinkOwner(source, sourcePath), bibliographyPath) + "#" + Uri.EscapeDataString(entryId));
        marker.SetAttributeValue("role", "doc-biblioref");
        AddSemanticToken(marker, "biblioref");
        var documents = new Dictionary<string, XDocument>(StringComparer.Ordinal) { [sourceManifestId] = source };
        if (backlinkText != null) {
            entry.Add(new XElement(Html + "p", new XElement(Html + "a", new XAttribute("role", "doc-backlink"),
                new XAttribute(Ops + "type", "backlink"), new XAttribute("href",
                    RelativeHref(HtmlContentLinkOwner(bibliography, bibliographyPath), sourcePath) + "#" + Uri.EscapeDataString(referenceId)), backlinkText)));
            documents[bibliographyManifestId] = bibliography;
        }
        CommitContentEdits(documents, cancellationToken);
    }

    private static XElement RequireBibliographyList(XElement list) {
        XElement? section = list.Parent;
        if ((list.Name != Html + "ol" && list.Name != Html + "ul") || section?.Name != Html + "section" ||
            !list.Ancestors(Html + "body").Any() || list.Elements().Any(element => element.Name != Html + "li"))
            throw new InvalidDataException("A bibliography requires an ol/ul directly inside a body section, containing list items.");
        return section;
    }
}
