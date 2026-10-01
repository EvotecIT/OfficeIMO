using OfficeIMO.Email.AddressBook;

namespace OfficeIMO.Email;

/// <summary>One provenance-tagged Outlook contact for offline consolidation review.</summary>
public sealed class EmailContactSource {
    /// <summary>Creates a source. IDs must be unique within one review.</summary>
    public EmailContactSource(string sourceId, EmailDocument document) {
        if (string.IsNullOrWhiteSpace(sourceId) || sourceId.Length > 1024) throw new ArgumentException("A source ID of at most 1024 characters is required.", nameof(sourceId));
        SourceId = sourceId; Document = document ?? throw new ArgumentNullException(nameof(document));
    }
    /// <summary>Caller-owned source identity, retained in every field choice.</summary>
    public string SourceId { get; }
    /// <summary>Contact model read during review. No input is changed.</summary>
    public EmailDocument Document { get; }
    /// <summary>Projects a person OAB entry through its canonical Outlook contact mapper.</summary>
    public static EmailContactSource FromAddressBook(OfflineAddressBookEntry entry, string? sourceId = null) {
        if (entry == null) throw new ArgumentNullException(nameof(entry));
        if (entry.IsDistributionList) throw new ArgumentException("Distribution lists require a group export.", nameof(entry));
        return new EmailContactSource(sourceId ?? entry.Reference.Id,
            new EmailDocument { OutlookItemKind = OutlookItemKind.Contact, Contact = entry.ToOutlookContact() });
    }
}

/// <summary>Bounded offline contact review policy.</summary>
public sealed class EmailContactReviewOptions {
    /// <summary>Creates limits for input count, aggregate serialized contact bytes and properties per card.</summary>
    public EmailContactReviewOptions(int maxContacts = 1000, long maxTotalBytes = 64L * 1024 * 1024,
        long maxContactBytes = 1024 * 1024, int maxPropertiesPerCard = 1024) {
        if (maxContacts <= 0 || maxContacts > 10000) throw new ArgumentOutOfRangeException(nameof(maxContacts));
        if (maxTotalBytes <= 0 || maxTotalBytes > 256L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maxTotalBytes));
        if (maxContactBytes <= 0 || maxContactBytes > maxTotalBytes) throw new ArgumentOutOfRangeException(nameof(maxContactBytes));
        if (maxPropertiesPerCard <= 0 || maxPropertiesPerCard > 10000) throw new ArgumentOutOfRangeException(nameof(maxPropertiesPerCard));
        MaxContacts = maxContacts; MaxTotalBytes = maxTotalBytes; MaxContactBytes = maxContactBytes; MaxPropertiesPerCard = maxPropertiesPerCard;
    }
    /// <summary>Maximum contact sources; exceeding this fails instead of returning an incomplete review.</summary>
    public int MaxContacts { get; }
    /// <summary>Maximum sum of serialized source cards.</summary>
    public long MaxTotalBytes { get; }
    /// <summary>Maximum serialized bytes of each source and consolidated output.</summary>
    public long MaxContactBytes { get; }
    /// <summary>Maximum properties in each retained card.</summary>
    public int MaxPropertiesPerCard { get; }
}

/// <summary>A portable field variant with source provenance.</summary>
public sealed class EmailContactFieldChoice {
    private readonly ContentLineProperty[] _properties;
    private ContentLineProperty Property => _properties[0];
    internal EmailContactFieldChoice(ContentLineProperty[] properties, IEnumerable<string> sourceIds) {
        _properties = properties; SourceIds = Array.AsReadOnly(sourceIds.ToArray());
    }
    /// <summary>Content-line property name.</summary>
    public string Name => Property.Name;
    /// <summary>Optional vCard property group.</summary>
    public string? Group => Property.Group;
    /// <summary>Ordered parameter descriptions, including repeated parameters, for reviewing each field's meaning.</summary>
    public IReadOnlyList<string> Parameters => Array.AsReadOnly(Property.Parameters
        .Select(parameter => parameter.Name + "=" + string.Join(",", parameter.Values)).ToArray());
    /// <summary>First raw format-escaped value in this variant.</summary>
    public string Value => Property.Value;
    /// <summary>All ordered values of this repeated property on the source card. Consolidation retains the selected set.</summary>
    public IReadOnlyList<string> Values => Array.AsReadOnly(_properties.Select(property => property.Value).ToArray());
    /// <summary>Source contacts carrying this exact field variant.</summary>
    public IReadOnlyList<string> SourceIds { get; }
    internal IEnumerable<ContentLineProperty> CopyProperties() {
        foreach (ContentLineProperty property in _properties) {
            var copy = new ContentLineProperty(property.Name, property.Value) { Group = property.Group };
            foreach (ContentLineParameter parameter in property.Parameters)
                copy.Parameters.Add(new ContentLineParameter(parameter.Name, parameter.Values.ToArray()));
            yield return copy;
        }
    }
}

/// <summary>Variants of one grouped, parameter-qualified contact property.</summary>
public sealed class EmailContactReviewField {
    internal EmailContactReviewField(string key, IReadOnlyList<EmailContactFieldChoice> choices) { Key = key; Choices = choices; }
    /// <summary>Opaque stable key used in explicit field selection.</summary>
    public string Key { get; }
    /// <summary>Different values, in deterministic source order.</summary>
    public IReadOnlyList<EmailContactFieldChoice> Choices { get; }
    /// <summary>Whether consolidation requires the caller to select a variant.</summary>
    public bool HasConflict => Choices.Count > 1;
}

/// <summary>A candidate contact group linked by exact SMTP or an unambiguous authoritative OAB identity.</summary>
public sealed class EmailContactReviewGroup {
    internal EmailContactReviewGroup(IReadOnlyList<string> ids, IReadOnlyList<EmailContactReviewField> fields) { SourceIds = ids; Fields = fields; }
    /// <summary>Contacts connected through an address identity. Shared addresses can belong to different people; grouping is a review suggestion.</summary>
    public IReadOnlyList<string> SourceIds { get; }
    /// <summary>Portable fields and all competing values.</summary>
    public IReadOnlyList<EmailContactReviewField> Fields { get; }
}

/// <summary>A read-only candidate review; it makes no changes to stores or address books.</summary>
public sealed class EmailContactReviewResult {
    internal EmailContactReviewResult(IReadOnlyList<EmailContactReviewGroup> groups, IReadOnlyList<EmailDiagnostic> diagnostics,
        long maxOutputBytes, int maxProperties) { Groups = groups; Diagnostics = diagnostics; MaxOutputBytes = maxOutputBytes; MaxProperties = maxProperties; }
    /// <summary>Candidate groups, including unlinked single contacts.</summary>
    public IReadOnlyList<EmailContactReviewGroup> Groups { get; }
    /// <summary>Semantic projection losses and directory ambiguity/incomplete coverage.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    internal long MaxOutputBytes { get; }
    internal int MaxProperties { get; }
}
