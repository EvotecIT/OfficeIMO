namespace OfficeIMO.OpenDocument;

/// <summary>Tracked paragraph change kinds authored by the native ODT surface.</summary>
public enum OdtTrackedChangeKind {
    /// <summary>Inserted content bounded by change-start and change-end markers.</summary>
    Insertion,
    /// <summary>Deleted content retained in the tracked-changes declaration.</summary>
    Deletion
}

/// <summary>An XML-backed ODT tracked change.</summary>
public sealed class OdtTrackedChange {
    private readonly OdtDocument _document;
    private readonly XElement _region;
    internal OdtTrackedChange(OdtDocument document, XElement region) { _document = document; _region = region; }
    /// <summary>Stable change identifier.</summary>
    public string Id => (string?)_region.Attribute(XNamespace.Xml + "id")
        ?? (string?)_region.Attribute(OdfNamespaces.Text + "id")
        ?? string.Empty;
    /// <summary>Change kind.</summary>
    public OdtTrackedChangeKind Kind => _region.Element(OdfNamespaces.Text + "deletion") != null
        ? OdtTrackedChangeKind.Deletion : OdtTrackedChangeKind.Insertion;
    /// <summary>Recorded creator.</summary>
    public string? Creator => ChangeInfo?.Element(OdfNamespaces.Dc + "creator")?.Value;
    /// <summary>Recorded change timestamp.</summary>
    public DateTimeOffset? Date {
        get {
            string? value = ChangeInfo?.Element(OdfNamespaces.Dc + "date")?.Value;
            return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTimeOffset parsed)
                ? parsed : (DateTimeOffset?)null;
        }
    }
    /// <summary>Plain deleted content for a deletion change, or empty text for an insertion declaration.</summary>
    public string DeletedText => Kind == OdtTrackedChangeKind.Deletion
        ? string.Join("\n", ChangeElement.Elements().Where(IsTextBlock).Select(OdfTextCodec.Read)) : string.Empty;
    /// <summary>Accepts this change in the owning document.</summary>
    public void Accept() => _document.AcceptTrackedChange(Id);
    /// <summary>Rejects this change in the owning document.</summary>
    public void Reject() => _document.RejectTrackedChange(Id);

    internal XElement Region => _region;
    internal XElement ChangeElement => _region.Elements().First(element =>
        element.Name == OdfNamespaces.Text + "insertion" || element.Name == OdfNamespaces.Text + "deletion");
    private XElement? ChangeInfo => ChangeElement.Element(OdfNamespaces.Office + "change-info");
    private static bool IsTextBlock(XElement element) => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h";
}

public sealed partial class OdtDocument {
    /// <summary>Tracked changes declared by the document.</summary>
    public IReadOnlyList<OdtTrackedChange> TrackedChanges => GetTrackedChangesContainer(create: false)?
        .Elements(OdfNamespaces.Text + "changed-region")
        .Select(element => new OdtTrackedChange(this, element)).ToList()
        ?? (IReadOnlyList<OdtTrackedChange>)Array.Empty<OdtTrackedChange>();

    /// <summary>Appends a tracked paragraph insertion.</summary>
    public OdtTrackedChange AddTrackedParagraphInsertion(string? text, string creator,
        DateTimeOffset? date = null) {
        ValidateCreator(creator);
        string id = NextTrackedChangeId();
        XElement region = CreateChangedRegion(id, OdfNamespaces.Text + "insertion", creator, date);
        GetTrackedChangesContainer(create: true)!.Add(region);
        var start = new XElement(OdfNamespaces.Text + "change-start", new XAttribute(OdfNamespaces.Text + "change-id", id));
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        var end = new XElement(OdfNamespaces.Text + "change-end", new XAttribute(OdfNamespaces.Text + "change-id", id));
        TextBody.Add(start, paragraph, end);
        MarkPartDirty("content.xml");
        return new OdtTrackedChange(this, region);
    }

    /// <summary>Marks an existing top-level body paragraph or heading as deleted while retaining its XML in change metadata.</summary>
    public OdtTrackedChange DeleteParagraphTracked(OdtParagraph paragraph, string creator,
        DateTimeOffset? date = null) {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        ValidateCreator(creator);
        XElement source = paragraph.Element;
        if (!ReferenceEquals(source.Parent, TextBody)) throw new ArgumentException("Paragraph must be a top-level block in this document body.", nameof(paragraph));
        string id = NextTrackedChangeId();
        XElement region = CreateChangedRegion(id, OdfNamespaces.Text + "deletion", creator, date);
        region.Element(OdfNamespaces.Text + "deletion")!.Add(new XElement(source));
        GetTrackedChangesContainer(create: true)!.Add(region);
        source.ReplaceWith(new XElement(OdfNamespaces.Text + "change", new XAttribute(OdfNamespaces.Text + "change-id", id)));
        MarkPartDirty("content.xml");
        return new OdtTrackedChange(this, region);
    }

    /// <summary>Accepts one tracked insertion or deletion.</summary>
    public bool AcceptTrackedChange(string id) {
        OdtTrackedChange? change = FindTrackedChange(id);
        if (change == null) return false;
        if (change.Kind == OdtTrackedChangeKind.Insertion) RemoveInsertionMarkers(id, removeContent: false);
        else RemoveDeletionMarker(id);
        change.Region.Remove();
        RemoveEmptyTrackedChangesContainer();
        MarkPartDirty("content.xml");
        return true;
    }

    /// <summary>Rejects one tracked insertion or deletion.</summary>
    public bool RejectTrackedChange(string id) {
        OdtTrackedChange? change = FindTrackedChange(id);
        if (change == null) return false;
        if (change.Kind == OdtTrackedChangeKind.Insertion) {
            RemoveInsertionMarkers(id, removeContent: true);
        } else {
            XElement? marker = TextBody.Descendants(OdfNamespaces.Text + "change")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OdfNamespaces.Text + "change-id"), id, StringComparison.Ordinal));
            if (marker != null) {
                List<XNode> restored = change.ChangeElement.Nodes()
                    .Where(node => !(node is XElement element && element.Name == OdfNamespaces.Office + "change-info"))
                    .Select(CloneNode).ToList();
                marker.ReplaceWith(restored);
            }
        }
        change.Region.Remove();
        RemoveEmptyTrackedChangesContainer();
        MarkPartDirty("content.xml");
        return true;
    }

    private OdtTrackedChange? FindTrackedChange(string id) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Change identifier cannot be empty.", nameof(id));
        return TrackedChanges.FirstOrDefault(change => string.Equals(change.Id, id, StringComparison.Ordinal));
    }

    private XElement? GetTrackedChangesContainer(bool create) {
        XElement? container = TextBody.Element(OdfNamespaces.Text + "tracked-changes");
        if (container == null && create) {
            container = new XElement(OdfNamespaces.Text + "tracked-changes");
            TextBody.AddFirst(container);
            MarkPartDirty("content.xml");
        }
        return container;
    }

    private string NextTrackedChangeId() {
        var ids = new HashSet<string>(TrackedChanges.Select(change => change.Id), StringComparer.Ordinal);
        int index = 1; string id;
        do { id = "ct" + index++.ToString(CultureInfo.InvariantCulture); } while (ids.Contains(id));
        return id;
    }

    private static XElement CreateChangedRegion(string id, XName changeName, string creator, DateTimeOffset? date) {
        return new XElement(OdfNamespaces.Text + "changed-region",
            new XAttribute(XNamespace.Xml + "id", id),
            new XAttribute(OdfNamespaces.Text + "id", id),
            new XElement(changeName,
                new XElement(OdfNamespaces.Office + "change-info",
                    new XElement(OdfNamespaces.Dc + "creator", creator),
                    new XElement(OdfNamespaces.Dc + "date", (date ?? DateTimeOffset.UtcNow).ToString("o", CultureInfo.InvariantCulture)))));
    }

    private void RemoveInsertionMarkers(string id, bool removeContent) {
        XElement? start = TextBody.Descendants(OdfNamespaces.Text + "change-start")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute(OdfNamespaces.Text + "change-id"), id, StringComparison.Ordinal));
        XElement? end = TextBody.Descendants(OdfNamespaces.Text + "change-end")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute(OdfNamespaces.Text + "change-id"), id, StringComparison.Ordinal));
        if (start == null || end == null || !ReferenceEquals(start.Parent, end.Parent)) { start?.Remove(); end?.Remove(); return; }
        if (removeContent) {
            XNode? node = start.NextNode;
            while (node != null && !ReferenceEquals(node, end)) { XNode? next = node.NextNode; node.Remove(); node = next; }
        }
        start.Remove(); end.Remove();
    }

    private void RemoveDeletionMarker(string id) {
        foreach (XElement marker in TextBody.Descendants(OdfNamespaces.Text + "change")
                     .Where(element => string.Equals((string?)element.Attribute(OdfNamespaces.Text + "change-id"), id, StringComparison.Ordinal)).ToList()) marker.Remove();
    }

    private void RemoveEmptyTrackedChangesContainer() {
        XElement? container = GetTrackedChangesContainer(create: false);
        if (container != null && !container.Elements(OdfNamespaces.Text + "changed-region").Any()) container.Remove();
    }

    private static XNode CloneNode(XNode node) => node is XElement element ? new XElement(element) :
        node is XText text ? new XText(text.Value) : node is XComment comment ? new XComment(comment.Value) : new XText(node.ToString());
    private static void ValidateCreator(string creator) {
        if (string.IsNullOrWhiteSpace(creator)) throw new ArgumentException("Tracked-change creator cannot be empty.", nameof(creator));
    }
}
