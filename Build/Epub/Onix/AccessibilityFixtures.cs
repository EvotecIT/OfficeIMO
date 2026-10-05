using OfficeIMO.Workflows;

internal static class AccessibilityFixtures {
    // Synthetic declarations exercise serialization only. These samples are not accessibility certifications.
    internal static BookOnixAccessibilityMetadata? Create(string profile) => profile switch {
        "accessibility-unknown" => new() {
            Status = BookOnixAccessibilityStatus.Unknown,
            Summary = "Accessibility has not been assessed. This is a schema qualification fixture.",
            PublisherInformationUrl = "https://example.org/accessibility",
            PublisherContactEmail = "accessibility@example.org"
        },
        "accessibility-claims" => new() {
            Summary = "Synthetic feature and conformance values for schema validation only; not an assessment of the accompanying EPUB.",
            Features = Enum.GetValues<BookOnixAccessibilityFeature>(),
            Conformance = new(BookOnixWcagVersion.V2_2, BookOnixWcagLevel.AA),
            AssessmentDate = new DateOnly(2026, 10, 5),
            PublisherInformationUrl = "https://example.org/accessibility?edition=1&format=epub",
            PublisherContactEmail = "accessibility@example.org"
        },
        _ => null
    };
}
