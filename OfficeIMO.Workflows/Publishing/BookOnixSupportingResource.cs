namespace OfficeIMO.Workflows;

/// <summary>An absolute HTTP(S) link to a resource, optionally tagged with an ONIX list 74 language. Never fetched.</summary>
public sealed record BookOnixResourceLink(string Url, string? LanguageCode = null);

/// <summary>Publisher-supplied promotional resource metadata, independent of resources embedded in the EPUB.</summary>
public sealed record BookOnixSupportingResource {
    /// <summary>Explicit purpose of the resource.</summary>
    public required BookOnixResourceContentType Type { get; init; }
    /// <summary>Explicit content mode, not inferred from the URL or file format.</summary>
    public required BookOnixResourceMode Mode { get; init; }
    /// <summary>One to 13 distinct recipients; unrestricted must stand alone.</summary>
    public required IReadOnlyList<BookOnixContentAudience> Audiences { get; init; }
    /// <summary>Optional territory where the resource may be used.</summary>
    public BookOnixTerritory? Territory { get; init; }
    /// <summary>Up to 16 plain-text translations of the required display credit.</summary>
    public IReadOnlyList<BookOnixCollateralTextValue> Credits { get; init; } = [];
    /// <summary>Up to 16 plain-text caption translations.</summary>
    public IReadOnlyList<BookOnixCollateralTextValue> Captions { get; init; } = [];
    /// <summary>Up to 16 plain-text translations identifying the copyright holder.</summary>
    public IReadOnlyList<BookOnixCollateralTextValue> CopyrightHolders { get; init; } = [];
    /// <summary>Up to 16 plain-text alternative-text translations. Presence does not establish accessibility.</summary>
    public IReadOnlyList<BookOnixCollateralTextValue> AlternativeTexts { get; init; } = [];
    /// <summary>
    /// Up to 16 distinct identity references, each matching a declared product contributor identity.
    /// Proprietary values must be unambiguous across schemes because resource features omit scheme names.
    /// Collection-only contributors do not satisfy these references.
    /// </summary>
    public IReadOnlyList<BookOnixContributorIdentifier> ContributorReferences { get; init; } = [];
    /// <summary>Optional nonnegative approximate audio/video duration in whole minutes.</summary>
    public int? LengthMinutes { get; init; }
    /// <summary>One to 16 explicit versions of the same resource.</summary>
    public required IReadOnlyList<BookOnixResourceVersion> Versions { get; init; }
}

/// <summary>A particular delivery version of a supporting resource. No files are fetched, measured, executed or deleted.</summary>
public sealed record BookOnixResourceVersion {
    /// <summary>Explicit hosting/delivery form.</summary>
    public required BookOnixResourceForm Form { get; init; }
    /// <summary>One to 16 distinct URL/language pairs. Multilingual links must all carry explicit languages; alternate URLs in one language are allowed.</summary>
    public required IReadOnlyList<BookOnixResourceLink> Links { get; init; }
    /// <summary>Optional ONIX list 178 code. Codes describe formats without introducing a runtime dependency.</summary>
    public string? FileFormatCode { get; init; }
    /// <summary>Optional positive image width in pixels.</summary>
    public int? ImageWidth { get; init; }
    /// <summary>Optional positive image height in pixels.</summary>
    public int? ImageHeight { get; init; }
    /// <summary>Optional display/download filename, not a path; at most 255 UTF-16 code units.</summary>
    public string? FileName { get; init; }
    /// <summary>Optional nonnegative exact file length in bytes, supplied by the publisher.</summary>
    public long? ByteLength { get; init; }
    /// <summary>Optional 64-hex-digit SHA-256 assertion, not computed or verified against a download.</summary>
    public string? Sha256 { get; init; }
    /// <summary>Up to 32 explicit version-specific usage constraints.</summary>
    public IReadOnlyList<BookOnixUsageConstraint> UsageConstraints { get; init; } = [];
    /// <summary>Up to 16 explicit version-specific licenses; caller-supplied schema validation still applies.</summary>
    public IReadOnlyList<BookOnixLicense> Licenses { get; init; } = [];
    /// <summary>Optional first permitted-use date; serialized but not enforced.</summary>
    public DateOnly? UsableFrom { get; init; }
    /// <summary>Optional last permitted-use date.</summary>
    public DateOnly? UsableUntil { get; init; }
    /// <summary>Optional resource-version update date.</summary>
    public DateOnly? UpdatedOn { get; init; }
}
