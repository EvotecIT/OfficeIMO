using System.Globalization;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>Appends an EPUB 3 title and its refinements atomically without changing the existing primary title.</summary>
    public void AddTitle(string id, EpubTitleMetadata title) {
        if (title == null) throw new ArgumentNullException(nameof(title));
        RequirePublishingMetadataId(id);
        List<XElement> records = TitleRecords(id, title);
        AppendPublishingMetadata(records);
    }

    /// <summary>
    /// Updates the first title and its EPUB 3 refinements atomically. An existing title id must match id;
    /// existing unrelated refinements and attributes are retained. Null optional fields retain their values.
    /// </summary>
    public void SetPrimaryTitle(string id, EpubTitleMetadata title) {
        if (title == null) throw new ArgumentNullException(nameof(title));
        if (PackageVersion != "3.0") throw new NotSupportedException("Refined publishing metadata requires EPUB 3.");
        XElement? primary = RequireSection("metadata").Element(Dc + "title");
        string? existingId = (string?)primary?.Attribute("id");
        if (existingId != null && existingId != id) throw new ArgumentException("Retain the existing primary title identifier: " + existingId, nameof(id));
        if (existingId == null) VerifyAvailableId(id);
        else XmlConvert.VerifyNCName(id);
        List<XElement> records = TitleRecords(id, title);
        EditPackageElement(RequireSection("metadata"), proposed => {
            XElement? target = proposed.Element(Dc + "title");
            if (target == null) { proposed.Add(records.Select(element => new XElement(element))); return; }
            target.Value = records[0].Value;
            target.SetAttributeValue("id", id);
            if (records[0].Attribute(XNamespace.Xml + "lang") != null)
                target.SetAttributeValue(XNamespace.Xml + "lang", (string?)records[0].Attribute(XNamespace.Xml + "lang"));
            foreach (XElement record in records.Skip(1)) {
                XElement? existing = proposed.Elements(Opf + "meta").FirstOrDefault(element =>
                    EpubVocabulary.Expand(proposed.Document!.Root!, (string?)element.Attribute("property") ?? string.Empty) ==
                    EpubVocabulary.Expand(proposed.Document!.Root!, (string)record.Attribute("property")!) &&
                    (string?)element.Attribute("refines") is string reference && ReferencesPackageId(reference, id));
                if (existing == null) proposed.Add(new XElement(record));
                else { existing.Value = record.Value; existing.SetAttributeValue("scheme", null); }
            }
        });
    }

    /// <summary>Appends an EPUB 3 subject with optional authority and code in one metadata edit.</summary>
    public void AddSubject(string id, EpubSubjectMetadata subject) {
        if (subject == null) throw new ArgumentNullException(nameof(subject));
        RequirePublishingMetadataId(id);
        ValidatePublishingName(subject.Text, null, subject.Language);
        if (subject.Authority != null) ValidatePublishingName(subject.Authority, null, null);
        if (subject.Code != null) {
            ValidatePublishingName(subject.Code, null, null);
            if (subject.Authority == null) throw new ArgumentException("A subject code requires its classification authority.", nameof(subject));
        }
        var value = new XElement(Dc + "subject", new XAttribute("id", id), subject.Text);
        value.SetAttributeValue(XNamespace.Xml + "lang", subject.Language);
        var records = new List<XElement> { value };
        if (subject.Authority != null) records.Add(PublishingRefinement(id, "authority", subject.Authority));
        if (subject.Code != null) records.Add(PublishingRefinement(id, "term", subject.Code));
        AppendPublishingMetadata(records);
    }

    /// <summary>
    /// Appends an EPUB 3 identifier with an ONIX type for ISBNs or DOI. ISBN shape and checksums are validated;
    /// range assignment and identifier registration are not. The selected package identity is unchanged.
    /// </summary>
    public void AddIdentifier(string id, EpubIdentifierMetadata identifier) {
        if (identifier == null) throw new ArgumentNullException(nameof(identifier));
        RequirePublishingMetadataId(id);
        string value = EpubBibliographicIdentifier.Normalize(identifier.Value, identifier.Kind);
        string? code = identifier.Kind switch {
            EpubIdentifierKind.Unspecified => null,
            EpubIdentifierKind.Isbn10 => "02",
            EpubIdentifierKind.Isbn13 => "15",
            EpubIdentifierKind.Doi => "06",
            _ => throw new ArgumentOutOfRangeException(nameof(identifier.Kind))
        };
        var records = new List<XElement> { new XElement(Dc + "identifier", new XAttribute("id", id), value) };
        if (code != null) {
            if (EpubVocabulary.Expand(Root, "onix:codelist5") != "http://www.editeur.org/ONIX/book/codelists/current.html#codelist5")
                throw new InvalidOperationException("The onix prefix does not identify the ONIX code list vocabulary.");
            XElement refinement = PublishingRefinement(id, "identifier-type", code);
            refinement.SetAttributeValue("scheme", "onix:codelist5");
            records.Add(refinement);
        }
        AppendPublishingMetadata(records);
    }

    /// <summary>
    /// Atomically updates supplied primary EPUB 3 publisher, description, rights and publication-date values.
    /// Null fields retain existing values. Other values, attributes and refinements remain intact.
    /// </summary>
    public void SetPublicationDetails(EpubPublicationDetails details) {
        if (details == null) throw new ArgumentNullException(nameof(details));
        if (PackageVersion != "3.0") throw new NotSupportedException("Typed publication details require EPUB 3.");
        var values = new[] {
            new KeyValuePair<string, string?>("publisher", details.Publisher),
            new KeyValuePair<string, string?>("description", details.Description),
            new KeyValuePair<string, string?>("rights", details.Rights),
            new KeyValuePair<string, string?>("date", details.PublicationDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
        }.Where(pair => pair.Value != null).ToArray();
        foreach (var pair in values) ValidatePublishingName(pair.Value!, null, null);
        if (values.Length == 0) return;
        EditPackageElement(RequireSection("metadata"), proposed => {
            foreach (var pair in values) {
                XElement? existing = proposed.Element(Dc + pair.Key);
                if (existing == null) proposed.Add(new XElement(Dc + pair.Key, pair.Value));
                else existing.Value = pair.Value!;
            }
        });
    }

    private static List<XElement> TitleRecords(string id, EpubTitleMetadata title) {
        ValidatePublishingName(title.Text, title.FileAs, title.Language);
        string kind = title.Kind switch {
            EpubTitleKind.Main => "main", EpubTitleKind.Subtitle => "subtitle", EpubTitleKind.Short => "short",
            EpubTitleKind.Collection => "collection", EpubTitleKind.Edition => "edition", EpubTitleKind.Expanded => "expanded",
            _ => throw new ArgumentOutOfRangeException(nameof(title.Kind))
        };
        var value = new XElement(Dc + "title", new XAttribute("id", id), title.Text);
        value.SetAttributeValue(XNamespace.Xml + "lang", title.Language);
        var records = new List<XElement> { value, PublishingRefinement(id, "title-type", kind) };
        if (title.DisplaySequence.HasValue) records.Add(PublishingRefinement(id, "display-seq", title.DisplaySequence.Value.ToString(CultureInfo.InvariantCulture)));
        if (title.FileAs != null) records.Add(PublishingRefinement(id, "file-as", title.FileAs));
        return records;
    }
}
