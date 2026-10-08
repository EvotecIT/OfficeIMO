namespace OfficeIMO.Workflows;

/// <summary>ONIX list 71 restrictions applicable to this digital-download export profile. Internal-only and print-on-demand codes are excluded.</summary>
public enum BookOnixSalesRestrictionKind {
    /// <summary>Restriction explained by required notes (00).</summary>
    Unspecified = 0,
    /// <summary>Retailer restriction where the more specific exclusive or own-brand distinction is unavailable (01).</summary>
    RetailerExclusiveOrOwnBrand = 1,
    /// <summary>Office-supplies channel only (02).</summary>
    OfficeSuppliesOnly = 2,
    /// <summary>Named retailers under the publisher brand (04).</summary>
    RetailerExclusive = 4,
    /// <summary>Named retailers under their own brand (05).</summary>
    RetailerOwnBrand = 5,
    /// <summary>Libraries only (06).</summary>
    LibrariesOnly = 6,
    /// <summary>Schools only (07).</summary>
    SchoolsOnly = 7,
    /// <summary>Publisher assertion of German indexed status (08).</summary>
    GermanIndexed = 8,
    /// <summary>Excludes library supply (09).</summary>
    ExceptLibraries = 9,
    /// <summary>News outlets only (10).</summary>
    NewsOutletsOnly = 10,
    /// <summary>Excludes named retailers (11).</summary>
    RetailerException = 11,
    /// <summary>Excludes subscription services (12).</summary>
    ExceptSubscriptionServices = 12,
    /// <summary>Subscription services only (13).</summary>
    SubscriptionServicesOnly = 13,
    /// <summary>Excludes online retail (14).</summary>
    ExceptOnlineRetail = 14,
    /// <summary>Online retail only (15).</summary>
    OnlineRetailOnly = 15,
    /// <summary>Excludes school supply (16).</summary>
    ExceptSchools = 16,
    /// <summary>Retail supply plus streaming through named subscription services (20).</summary>
    SelectedSubscriptionServices = 20,
    /// <summary>Named subscription services only (21).</summary>
    SubscriptionServiceExclusive = 21,
    /// <summary>Educational institutions only (22).</summary>
    EducationOnly = 22,
    /// <summary>Excludes educational institutions (23).</summary>
    ExceptEducation = 23,
    /// <summary>Explicit unrestricted sales assertion (99).</summary>
    NoRestrictions = 99,
}

/// <summary>ONIX list 102 outlet identifier schemes.</summary>
public enum BookOnixSalesOutletScheme {
    /// <summary>Proprietary identifier, requiring a scheme name (01).</summary>
    Proprietary = 1,
    /// <summary>ONIX list 139 outlet code; lexical check only, no registry lookup (03).</summary>
    Onix = 3,
    /// <summary>13-digit GLN; lexical check only, no checksum or assignment verification (04).</summary>
    Gln = 4,
    /// <summary>Seven-digit SAN; lexical check only, no checksum or assignment verification (05).</summary>
    San = 5
}

/// <summary>An explicit outlet identifier, at most 100 UTF-16 code units, preserved without normalization.</summary>
public sealed record BookOnixSalesOutletIdentifier(BookOnixSalesOutletScheme Scheme, string Value) {
    /// <summary>Required only for proprietary identifiers, at most 100 UTF-16 code units.</summary>
    public string? SchemeName { get; init; }
}

/// <summary>A named or identified outlet. At least one name or identifier is required.</summary>
public sealed record BookOnixSalesOutlet {
    /// <summary>Optional plain-text name, at most 200 UTF-16 code units.</summary>
    public string? Name { get; init; }
    /// <summary>Optional three-letter ONIX language code for Name; requires Name.</summary>
    public string? NameLanguageCode { get; init; }
    /// <summary>Up to eight identifiers, with distinct scheme and proprietary-name pairs.</summary>
    public IReadOnlyList<BookOnixSalesOutletIdentifier> Identifiers { get; init; } = [];
}

/// <summary>A sales restriction note, optionally tagged with an ONIX language code.</summary>
public sealed record BookOnixSalesRestrictionNote(string Text, string? LanguageCode = null) {
    /// <summary>Plain text by default; XHTML uses the shared bounded ONIX fragment profile. Text is limited to 300 UTF-16 code units after decoding, and XHTML source to 4096.</summary>
    public BookOnixCollateralTextFormat Format { get; init; } = BookOnixCollateralTextFormat.PlainText;
}

/// <summary>Explicit non-territorial restriction within the containing rights territory or supply market. Does not enforce purchase eligibility.</summary>
public sealed record BookOnixSalesRestriction(BookOnixSalesRestrictionKind Kind) {
    /// <summary>Up to 16 affected outlets. Required for retailer and selected-subscription restrictions.</summary>
    public IReadOnlyList<BookOnixSalesOutlet> Outlets { get; init; } = [];
    /// <summary>Up to 16 distinct-language translations, at most 300 decoded UTF-16 code units each (4096 source units for XHTML). Required for Unspecified.</summary>
    public IReadOnlyList<BookOnixSalesRestrictionNote> Notes { get; init; } = [];
    /// <summary>Optional first effective date.</summary>
    public DateOnly? ValidFrom { get; init; }
    /// <summary>Optional last effective date; cannot precede ValidFrom.</summary>
    public DateOnly? ValidUntil { get; init; }
}
