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
        "accessibility-provenance" => new() {
            Summary = "Synthetic certification provenance for schema validation; no organization or assessment is real.",
            Certification = new("Example Testing & Certification", "https://certifier.example.org/") {
                CredentiallingOrganizationName = "Example Credentials",
                CredentiallingOrganizationUrl = "https://credentials.example.org/"
            },
            IndependentReportUrl = "https://certifier.example.org/reports/edition-1",
            IntermediaryInformationUrl = "https://intermediary.example.org/edition-1",
            CompatibilityReportUrl = "https://example.org/compatibility/edition-1",
            IntermediaryContactEmail = "accessibility@intermediary.example.org"
        },
        _ => null
    };
}
