using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Appends an index term and up to 1024 labelled links to existing XHTML spine content.
    /// The destination ul must be directly inside a body section, or inside an existing index li for subentries.
    /// Optional subentriesId creates a nested ul for subsequent calls. Publisher order and labels are preserved;
    /// sorting, term extraction and page-number calculation are not inferred. The edit commits atomically.
    /// </summary>
    public void AddIndexEntry(string manifestId, string listId, string entryId, string term,
        IReadOnlyList<EpubIndexLocator> locators, string? subentriesId = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Semantic index authoring requires EPUB 3.");
        RequireText(listId, nameof(listId)); RequireText(term, nameof(term)); XmlConvert.VerifyNCName(entryId);
        if (locators == null) throw new ArgumentNullException(nameof(locators));
        if (locators.Count > 1024) throw new ArgumentOutOfRangeException(nameof(locators), "An index entry supports at most 1024 locators.");
        if (subentriesId != null) XmlConvert.VerifyNCName(subentriesId);
        if (locators.Count == 0 && subentriesId == null) throw new ArgumentException("An index entry requires locators or a subentry list.", nameof(locators));
        XDocument content = EditableXhtml(manifestId);
        XElement list = RequireContentElement(content, listId);
        if (list.Name != Html + "ul" || !list.Ancestors(Html + "body").Any() || list.Elements().Any(element => element.Name != Html + "li"))
            throw new InvalidDataException("An index destination requires a body ul containing list items.");
        XElement? section;
        if (list.Parent?.Name == Html + "section") section = list.Parent;
        else if (list.Parent?.Name == Html + "li") {
            section = list.Ancestors(Html + "section").FirstOrDefault(element =>
                HasToken((string?)element.Attribute(Ops + "type"), "index") && (string?)element.Attribute("role") == "doc-index");
            if (section == null) throw new InvalidDataException("A subentry list requires an existing semantic index section.");
        } else throw new InvalidDataException("An index list must be directly inside a section or an index entry.");
        RequireCompatibleRole(section, "doc-index");
        string owner = HtmlContentLinkOwner(content, RequireLocalPath(RequireManifestItem(manifestId)));
        var documents = new Dictionary<string, XDocument>(StringComparer.Ordinal) { [manifestId] = content };
        var entry = new XElement(Html + "li", new XAttribute("id", entryId), new XElement(Html + "span", term));
        for (int index = 0; index < locators.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            EpubIndexLocator locator = locators[index] ?? throw new ArgumentException("Index locators cannot contain null.", nameof(locators));
            RequireText(locator.ManifestId, nameof(locator.ManifestId)); RequireText(locator.Label, nameof(locator.Label));
            if (!Spine.Any(item => item.ManifestId == locator.ManifestId)) throw new InvalidDataException("Index locators must target spine documents.");
            if (!documents.TryGetValue(locator.ManifestId, out XDocument? target)) {
                target = EditableXhtml(locator.ManifestId); documents.Add(locator.ManifestId, target);
            }
            string path = RequireLocalPath(RequireManifestItem(locator.ManifestId));
            string href = RelativeHref(owner, path);
            if (locator.FragmentId != null) {
                RequireText(locator.FragmentId, nameof(locator.FragmentId));
                XElement destination = RequireContentElement(target, locator.FragmentId);
                if (!destination.AncestorsAndSelf(Html + "body").Any()) throw new InvalidDataException("An index fragment must target document body content.");
                href += "#" + Uri.EscapeDataString(locator.FragmentId);
            }
            entry.Add(index == 0 ? ", " : "; ", new XElement(Html + "a", new XAttribute("href", href), locator.Label));
        }
        if (subentriesId != null) entry.Add(new XElement(Html + "ul", new XAttribute("id", subentriesId)));
        list.Add(entry);
        AddSemanticToken(section, "index"); section.SetAttributeValue("role", "doc-index");
        CommitContentEdits(new Dictionary<string, XDocument>(StringComparer.Ordinal) { [manifestId] = content }, cancellationToken);
    }
}
