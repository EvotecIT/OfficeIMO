using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Appends a term and XHTML definition to an existing dl directly inside a section in the document body.
    /// Adds glossary semantics while preserving the section heading and existing entries. Term text is escaped;
    /// definition links use the receiving document's effective HTML base. Entries remain in caller-supplied order.
    /// </summary>
    public void AddGlossaryEntry(string manifestId, string listId, string termId, string term, string definitionXhtml,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Semantic glossary authoring requires EPUB 3.");
        RequireText(listId, nameof(listId)); RequireText(term, nameof(term)); RequireText(definitionXhtml, nameof(definitionXhtml));
        XmlConvert.VerifyNCName(termId);
        XDocument content = EditableXhtml(manifestId);
        XElement list = RequireContentElement(content, listId);
        XElement section = RequireGlossaryList(list);
        RequireCompatibleRole(section, "doc-glossary");
        XElement definition = new XElement(Html + "dd", ParseAuthoringFragment(definitionXhtml));
        list.Add(new XElement(Html + "dt", new XAttribute("id", termId), new XElement(Html + "dfn", term)), definition);
        AddSemanticToken(section, "glossary");
        section.SetAttributeValue("role", "doc-glossary");
        CommitContentEdits(new Dictionary<string, XDocument>(StringComparer.Ordinal) { [manifestId] = content }, cancellationToken);
    }

    /// <summary>
    /// Links an existing labelled anchor without an href to a glossary dt followed by one dd.
    /// An optional localized backlink is appended to that definition. Both documents are committed atomically;
    /// call once per occurrence to link multiple passages to one term. This does not request a reader popup.
    /// </summary>
    public void LinkGlossaryTerm(string sourceManifestId, string referenceId, string glossaryManifestId, string termId,
        string? backlinkText = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Semantic glossary authoring requires EPUB 3.");
        RequireText(referenceId, nameof(referenceId)); RequireText(termId, nameof(termId));
        if (backlinkText != null) RequireText(backlinkText, nameof(backlinkText));
        XDocument source = EditableXhtml(sourceManifestId);
        XDocument glossary = sourceManifestId == glossaryManifestId ? source : EditableXhtml(glossaryManifestId);
        XElement marker = RequireContentElement(source, referenceId);
        if (marker.Name != Html + "a" || marker.Attribute("href") != null || marker.Ancestors(Html + "a").Any() ||
            marker.Descendants(Html + "a").Any() || !marker.Ancestors(Html + "body").Any())
            throw new InvalidDataException("A glossary reference requires a body anchor without an href or nested anchor.");
        if (string.IsNullOrWhiteSpace(marker.Value) && string.IsNullOrWhiteSpace((string?)marker.Attribute("aria-label")))
            throw new InvalidDataException("A glossary reference needs visible text or an aria-label.");
        RequireCompatibleRole(marker, "doc-glossref");
        XElement term = RequireContentElement(glossary, termId);
        XElement? definition = term.ElementsAfterSelf().FirstOrDefault();
        if (term.Name != Html + "dt" || term.Parent == null || definition?.Name != Html + "dd" ||
            definition.ElementsAfterSelf().FirstOrDefault()?.Name == Html + "dd")
            throw new InvalidDataException("A glossary target requires a dt followed by exactly one dd.");
        XElement section = RequireGlossaryList(term.Parent);
        if (!HasToken((string?)section.Attribute("role"), "doc-glossary") || !HasToken((string?)section.Attribute(Ops + "type"), "glossary"))
            throw new InvalidDataException("The target must belong to a semantic glossary section and definition list.");
        string sourcePath = RequireLocalPath(RequireManifestItem(sourceManifestId));
        string glossaryPath = RequireLocalPath(RequireManifestItem(glossaryManifestId));
        marker.SetAttributeValue("href", RelativeHref(HtmlContentLinkOwner(source, sourcePath), glossaryPath) + "#" + Uri.EscapeDataString(termId));
        marker.SetAttributeValue("role", "doc-glossref");
        AddSemanticToken(marker, "glossref");
        var documents = new Dictionary<string, XDocument>(StringComparer.Ordinal) { [sourceManifestId] = source };
        if (backlinkText != null) {
            definition.Add(new XElement(Html + "p", new XElement(Html + "a", new XAttribute("role", "doc-backlink"),
                new XAttribute(Ops + "type", "backlink"), new XAttribute("href",
                    RelativeHref(HtmlContentLinkOwner(glossary, glossaryPath), sourcePath) + "#" + Uri.EscapeDataString(referenceId)), backlinkText)));
            documents[glossaryManifestId] = glossary;
        }
        CommitContentEdits(documents, cancellationToken);
    }

    private static XElement RequireGlossaryList(XElement list) {
        XElement? section = list.Parent;
        XElement[] children = list.Elements().ToArray();
        if (list.Name != Html + "dl" || section?.Name != Html + "section" || !list.Ancestors(Html + "body").Any() ||
            children.Any(element => element.Name != Html + "dt" && element.Name != Html + "dd") ||
            (children.Length > 0 && (children[0].Name != Html + "dt" || children[children.Length - 1].Name != Html + "dd")))
            throw new InvalidDataException("A glossary requires a dl directly inside a body section, containing complete dt/dd groups.");
        return section;
    }
}
