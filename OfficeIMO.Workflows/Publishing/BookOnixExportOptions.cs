namespace OfficeIMO.Workflows;

/// <summary>Complete-record notification types supported by the bibliographic ONIX export profile.</summary>
public enum BookOnixNotification {
    /// <summary>Early notification (ONIX 01).</summary>
    Early,
    /// <summary>Advance notification with confirmed information (ONIX 02).</summary>
    Advance,
    /// <summary>Complete record confirmed at or after publication (ONIX 03).</summary>
    Confirmed
}

/// <summary>Contributor roles supported by the ONIX export profile.</summary>
public enum BookOnixContributorRole {
    /// <summary>Author of the text (A01).</summary>
    Author,
    /// <summary>Editor (B01).</summary>
    Editor,
    /// <summary>Translator (B06).</summary>
    Translator,
    /// <summary>Illustrator (A12).</summary>
    Illustrator,
    /// <summary>Other creative responsibility (Z99).</summary>
    Other,
    /// <summary>Editor of the series to which the product belongs (B09).</summary>
    SeriesEditor
}

/// <summary>An explicitly classified ONIX credit; names are not parsed or inferred from EPUB creator text.</summary>
/// <param name="Name">Full credited name.</param>
/// <param name="Role">Creative responsibility.</param>
/// <param name="IsOrganization">Writes CorporateName instead of PersonName when true.</param>
public sealed record BookOnixContributor(string Name, BookOnixContributorRole Role, bool IsOrganization = false);

/// <summary>
/// Publisher-supplied assertions for a single-product ONIX 3.1 bibliographic record.
/// Only the selected title (the first by default) and selected ISBN are taken from the EPUB. Other EPUB metadata is not projected.
/// Commercial metadata is explicit and optional. This profile does not perform retailer submission.
/// </summary>
public sealed record BookOnixExportOptions {
    /// <summary>Organization sending the ONIX message.</summary>
    public required string SenderName { get; init; }
    /// <summary>Stable sender-owned record reference, independent of transmission time.</summary>
    public required string RecordReference { get; init; }
    /// <summary>Explicit message timestamp, serialized in UTC.</summary>
    public required DateTimeOffset SentAt { get; init; }
    /// <summary>Complete-record notification intent; block updates and deletion are not supported.</summary>
    public required BookOnixNotification Notification { get; init; }
    /// <summary>OPF id of the dc:identifier containing this digital edition's ISBN-13.</summary>
    public required string IdentifierId { get; init; }
    /// <summary>Optional OPF id of the title to export. When omitted, exports the first dc:title as before.</summary>
    public string? TitleId { get; init; }
    /// <summary>ONIX list 74 language code, such as eng, pol or fre. The supplied schema checks membership.</summary>
    public required string LanguageCode { get; init; }
    /// <summary>Publisher of this edition; not inferred from arbitrary Dublin Core dates or rights statements.</summary>
    public required string PublisherName { get; init; }
    /// <summary>Optional explicit sorting prefix or no-prefix assertion; omission retains unsplit TitleText.</summary>
    public BookOnixTitleSorting? TitleSorting { get; init; }
    /// <summary>Optional subtitle for the ONIX product-level title.</summary>
    public string? Subtitle { get; init; }
    /// <summary>At most 32 explicit alternative product-level titles, retained in caller order after the distinctive title.</summary>
    public IReadOnlyList<BookOnixAlternativeTitle> AlternativeTitles { get; init; } = [];
    /// <summary>Explicit publication date, when known.</summary>
    public DateOnly? PublicationDate { get; init; }
    /// <summary>Ordered credits, at most 100. Mutually exclusive with NoContributors.</summary>
    public IReadOnlyList<BookOnixContributor> Contributors { get; init; } = [];
    /// <summary>Explicit assertion that the product has no credited contributors. Missing EPUB credits do not imply this.</summary>
    public bool NoContributors { get; init; }
    /// <summary>At most 32 explicitly described collection memberships.</summary>
    public IReadOnlyList<BookOnixCollection> Collections { get; init; } = [];
    /// <summary>Explicit assertion of no collection membership, mutually exclusive with Collections.</summary>
    public bool NoCollection { get; init; }
    /// <summary>Optional explicit edition characteristics, numbering and multilingual statements.</summary>
    public BookOnixEdition? Edition { get; init; }
    /// <summary>At most 64 explicit subject declarations; free-form EPUB subject authorities are not mapped automatically.</summary>
    public IReadOnlyList<BookOnixSubject> Subjects { get; init; } = [];
    /// <summary>Optional explicit audience categories, age ranges and multilingual descriptions.</summary>
    public BookOnixAudienceMetadata? Audience { get; init; }
    /// <summary>At most 64 explicit collateral text items; recipient and usage declarations do not enforce access control.</summary>
    public IReadOnlyList<BookOnixCollateralText> CollateralTexts { get; init; } = [];
    /// <summary>Optional explicit publishing status, territorial rights, supplier availability and pricing.</summary>
    public BookOnixCommercialMetadata? Commercial { get; init; }
    /// <summary>Optional explicit accessibility assertions; never inferred from EPUB metadata or automated checks.</summary>
    public BookOnixAccessibilityMetadata? Accessibility { get; init; }
}
