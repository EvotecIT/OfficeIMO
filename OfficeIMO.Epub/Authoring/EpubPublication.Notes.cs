using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Adds an EPUB 3 note and links an existing labelled anchor marker to it, with a return link.
    /// Both content documents are committed atomically, including identifier and retention checks.
    /// Existing reference labels, document heads, and unrelated content are retained. Full publication
    /// validation remains part of Write/Preflight; this operation does not certify accessibility.
    /// </summary>
    public void AddNote(EpubNoteOptions options, CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Semantic note authoring requires EPUB 3.");
        RequireText(options.SourceManifestId, nameof(options.SourceManifestId));
        RequireText(options.NotesManifestId, nameof(options.NotesManifestId));
        RequireText(options.ReferenceId, nameof(options.ReferenceId));
        RequireText(options.ContainerId, nameof(options.ContainerId));
        RequireText(options.BacklinkText, nameof(options.BacklinkText));
        RequireText(options.BodyXhtml, nameof(options.BodyXhtml));
        XmlConvert.VerifyNCName(options.NoteId);
        if (!Enum.IsDefined(typeof(EpubNoteKind), options.Kind)) throw new ArgumentOutOfRangeException(nameof(options.Kind));
        XDocument source = EditableXhtml(options.SourceManifestId);
        XDocument notes = options.SourceManifestId == options.NotesManifestId ? source : EditableXhtml(options.NotesManifestId);
        string sourcePath = RequireLocalPath(RequireManifestItem(options.SourceManifestId));
        string notesPath = RequireLocalPath(RequireManifestItem(options.NotesManifestId));
        var documents = new Dictionary<string, XDocument>(StringComparer.Ordinal) { [options.SourceManifestId] = source };
        documents[options.NotesManifestId] = notes;
        XElement marker = RequireContentElement(source, options.ReferenceId);
        if (marker.Name != Html + "a" || marker.Attribute("href") != null || marker.Ancestors(Html + "a").Any())
            throw new InvalidDataException("A note reference must select an XHTML anchor without an existing href or enclosing anchor.");
        if (string.IsNullOrWhiteSpace(marker.Value) && string.IsNullOrWhiteSpace((string?)marker.Attribute("aria-label")))
            throw new InvalidDataException("A note reference needs visible text or an aria-label.");
        RequireCompatibleRole(marker, "doc-noteref");
        XElement container = RequireContentElement(notes, options.ContainerId);
        XElement note;
        if (options.Kind == EpubNoteKind.Footnote) {
            if (container.Name != Html + "section" && container.Name != Html + "div" && container.Name != Html + "body")
                throw new InvalidDataException("Footnotes require a section, div, or body container.");
            if (container.AncestorsAndSelf().Any(element => HasToken((string?)element.Attribute("role"), "doc-endnotes") ||
                HasToken((string?)element.Attribute(Ops + "type"), "endnotes")))
                throw new InvalidDataException("Use endnote semantics inside an endnotes section.");
            note = new XElement(Html + "aside", new XAttribute("role", "doc-footnote"), new XAttribute(Ops + "type", "footnote"));
        } else {
            if (container.Name != Html + "ol" && container.Name != Html + "ul")
                throw new InvalidDataException("Endnotes require an ordered or unordered list container.");
            XElement section = container.Ancestors(Html + "section").FirstOrDefault()
                ?? throw new InvalidDataException("An endnotes list must be inside an XHTML section.");
            RequireCompatibleRole(section, "doc-endnotes");
            AddSemanticToken(section, "endnotes");
            section.SetAttributeValue("role", "doc-endnotes");
            note = new XElement(Html + "li", new XAttribute(Ops + "type", "endnote"));
        }
        note.SetAttributeValue("id", options.NoteId);
        note.Add(ParseAuthoringFragment(options.BodyXhtml));
        string sourceOwner = HtmlContentLinkOwner(source, sourcePath);
        string notesOwner = HtmlContentLinkOwner(notes, notesPath);
        marker.SetAttributeValue("href", RelativeHref(sourceOwner, notesPath) + "#" + Uri.EscapeDataString(options.NoteId));
        marker.SetAttributeValue("role", "doc-noteref");
        AddSemanticToken(marker, "noteref");
        note.Add(new XElement(Html + "p", new XElement(Html + "a", new XAttribute("role", "doc-backlink"),
            new XAttribute("href", RelativeHref(notesOwner, sourcePath) + "#" + Uri.EscapeDataString(options.ReferenceId)), options.BacklinkText)));
        container.Add(note);
        CommitContentEdits(documents, cancellationToken);
    }

    private static void RequireCompatibleRole(XElement element, string role) {
        string? existing = (string?)element.Attribute("role");
        if (existing != null && existing != role) throw new InvalidDataException("Semantic authoring cannot replace an existing role: " + existing);
    }
    private static void AddSemanticToken(XElement element, string token) => element.SetAttributeValue(Ops + "type",
        string.Join(" ", Tokens((string?)element.Attribute(Ops + "type")).Concat(new[] { token }).Distinct(StringComparer.Ordinal)));
}
