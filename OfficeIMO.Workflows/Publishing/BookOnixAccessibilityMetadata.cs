namespace OfficeIMO.Workflows;

/// <summary>Publisher assertions from ONIX list 196; these are not inferred from individual EPUB elements.</summary>
public enum BookOnixAccessibilityFeature {
    /// <summary>Navigation reaches all structural levels and non-text items (11).</summary>
    TableOfContentsNavigation = 11,
    /// <summary>Index entries link to their occurrences (12).</summary>
    IndexNavigation = 12,
    /// <summary>Substantially all content participates in a logical reading sequence (13).</summary>
    SingleLogicalReadingOrder = 13,
    /// <summary>Substantially all meaningful non-text content has brief text alternatives (14).</summary>
    ShortTextAlternatives = 14,
    /// <summary>Substantially all meaningful non-text content has extended text alternatives (15).</summary>
    ExtendedTextAlternatives = 15,
    /// <summary>Substantially all text, including alternatives, has synchronized recorded narration (20).</summary>
    SynchronizedRecordedAudio = 20,
    /// <summary>Document and passage languages are identified where appropriate (22).</summary>
    LanguageTagging = 22,
    /// <summary>Markup supports forward/backward structural navigation at every level (29).</summary>
    StructuralNavigation = 29,
    /// <summary>ARIA semantics improve structure or navigation (30).</summary>
    AriaRoles = 30,
    /// <summary>Basic landmark navigation is present (32).</summary>
    LandmarkNavigation = 32,
    /// <summary>All text permits reader-controlled appearance changes and reflow (36).</summary>
    ModifiableTextAppearance = 36,
    /// <summary>Accessible explanations are provided for unusual terminology (38).</summary>
    ExplainedTerminology = 38,
    /// <summary>Link and control purposes are conveyed by text or separate descriptions (40).</summary>
    ClearLinkPurposes = 40,
    /// <summary>Navigation follows static print-equivalent or digital page boundaries (41).</summary>
    PageListNavigation = 41
}

/// <summary>Explicit limitation statements, not conformance levels.</summary>
public enum BookOnixAccessibilityStatus {
    /// <summary>Assessment or information is insufficient (08).</summary>
    Unknown,
    /// <summary>Significant accessibility limitations are known (09).</summary>
    Limited
}

/// <summary>WCAG versions supported by the EPUB Accessibility 1.1 declaration profile.</summary>
public enum BookOnixWcagVersion {
    /// <summary>WCAG 2.0 (80).</summary>
    V2_0,
    /// <summary>WCAG 2.1 (81).</summary>
    V2_1,
    /// <summary>WCAG 2.2 (82).</summary>
    V2_2
}

/// <summary>Declared WCAG conformance level.</summary>
public enum BookOnixWcagLevel {
    /// <summary>Level A (84).</summary>
    A,
    /// <summary>Level AA (85).</summary>
    AA,
    /// <summary>Level AAA (86).</summary>
    AAA
}

/// <summary>Explicit publisher assertion of EPUB Accessibility 1.1 plus a WCAG version and level; not a library certification.</summary>
public sealed record BookOnixEpubAccessibility11Conformance(BookOnixWcagVersion Version, BookOnixWcagLevel Level);

/// <summary>Optional publisher-supplied ONIX accessibility details. No fields are inferred from EPUB metadata or validation results.</summary>
public sealed record BookOnixAccessibilityMetadata {
    /// <summary>Summary of features and limitations (00), at most 4096 characters.</summary>
    public string? Summary { get; init; }
    /// <summary>Optional unknown or limited accessibility statement (08/09).</summary>
    public BookOnixAccessibilityStatus? Status { get; init; }
    /// <summary>Distinct explicit feature assertions, at most 32.</summary>
    public IReadOnlyList<BookOnixAccessibilityFeature> Features { get; init; } = [];
    /// <summary>Optional explicit EPUB Accessibility 1.1 and WCAG conformance assertion (04 plus version and level).</summary>
    public BookOnixEpubAccessibility11Conformance? Conformance { get; init; }
    /// <summary>Optional caller-supplied certifier identity and credentialling organization. Does not itself assert conformance.</summary>
    public BookOnixAccessibilityCertification? Certification { get; init; }
    /// <summary>HTTP(S) product-specific report maintained by an independent compliance or testing organization (94).</summary>
    public string? IndependentReportUrl { get; init; }
    /// <summary>HTTP(S) product-specific accessibility information maintained by a publisher-nominated trusted intermediary (95).</summary>
    public string? IntermediaryInformationUrl { get; init; }
    /// <summary>HTTP(S) description of compatibility testing, including assistive technology (97).</summary>
    public string? CompatibilityReportUrl { get; init; }
    /// <summary>Plain trusted intermediary contact email for accessibility questions (98).</summary>
    public string? IntermediaryContactEmail { get; init; }
    /// <summary>Date of the latest actual assessment (91); not the export date.</summary>
    public DateOnly? AssessmentDate { get; init; }
    /// <summary>Absolute HTTP(S) publisher-maintained accessibility information page (96). Never fetched by export.</summary>
    public string? PublisherInformationUrl { get; init; }
    /// <summary>Plain publisher contact email address for accessibility questions (99).</summary>
    public string? PublisherContactEmail { get; init; }
}
