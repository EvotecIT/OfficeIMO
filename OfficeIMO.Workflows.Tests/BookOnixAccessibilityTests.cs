using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixAccessibilityTests {
    private static readonly XNamespace Onix = BookProject.OnixNamespace;

    [Fact]
    public void ExplicitAssertionsUseAccessibilityCompositesAndLeaveTheProjectUnchanged() {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        var metadata = new BookOnixAccessibilityMetadata {
            Summary = "Text & navigation; images have known limitations.", Status = BookOnixAccessibilityStatus.Limited,
            Features = [BookOnixAccessibilityFeature.LanguageTagging, BookOnixAccessibilityFeature.PageListNavigation],
            AssessmentDate = new DateOnly(2026, 10, 5), PublisherInformationUrl = "https://example.org/accessibility?a=1&b=2",
            PublisherContactEmail = "accessibility@example.org"
        };
        var xml = Export(metadata); var features = xml.Descendants(Onix + "ProductFormFeature").ToArray();
        Assert.Equal(new[] { "00", "09", "22", "41", "91", "96", "99" }, features.Select(e => e.Element(Onix + "ProductFormFeatureValue")!.Value));
        Assert.All(features, e => Assert.Equal("09", e.Element(Onix + "ProductFormFeatureType")!.Value));
        Assert.Equal(metadata.Summary, Description(xml, "00"));
        Assert.Equal("20261005", Description(xml, "91"));
        Assert.Equal(metadata.PublisherInformationUrl, Description(xml, "96"));
        Assert.Equal(metadata.PublisherContactEmail, Description(xml, "99"));
        var result = project.ExportOnix(BookOnixTests.Options() with { Accessibility = metadata }, BookOnixTests.TestSchema());
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(project.Export().Bytes, result.Publication.Bytes);
    }

    [Theory]
    [InlineData(BookOnixWcagVersion.V2_0, BookOnixWcagLevel.A, "80", "84")]
    [InlineData(BookOnixWcagVersion.V2_1, BookOnixWcagLevel.AA, "81", "85")]
    [InlineData(BookOnixWcagVersion.V2_2, BookOnixWcagLevel.AAA, "82", "86")]
    public void ConformanceRequiresAnExplicitVersionAndLevel(BookOnixWcagVersion version, BookOnixWcagLevel level, string versionCode, string levelCode) {
        var xml = Export(new() { Conformance = new(version, level) });
        Assert.Equal(new[] { "04", versionCode, levelCode }, xml.Descendants(Onix + "ProductFormFeatureValue").Select(e => e.Value));
    }

    [Fact]
    public void EpubAccessibilityMetadataNeverBecomesAnImplicitOnixAssertion() {
        var project = BookOnixTests.Project();
        project.Publication.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["tableOfContents"], Hazards = ["none"], Summary = "Caller-provided EPUB description."
        });
        var xml = XDocument.Load(new MemoryStream(project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema()).Bytes));
        Assert.Empty(xml.Descendants(Onix + "ProductFormFeature"));
        Assert.Equal("08", Export(new() { Status = BookOnixAccessibilityStatus.Unknown }).Descendants(Onix + "ProductFormFeatureValue").Single().Value);
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("unknown-conformance")]
    [InlineData("duplicate")]
    [InlineData("feature")]
    [InlineData("status")]
    [InlineData("version")]
    [InlineData("level")]
    [InlineData("url")]
    [InlineData("credentials")]
    [InlineData("email")]
    [InlineData("oversized")]
    public void InvalidDeclarationsCannotChangeTheProject(string kind) {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        BookOnixAccessibilityMetadata metadata = kind switch {
            "unknown-conformance" => new() { Status = BookOnixAccessibilityStatus.Unknown, Conformance = new(BookOnixWcagVersion.V2_2, BookOnixWcagLevel.AA) },
            "duplicate" => new() { Features = [BookOnixAccessibilityFeature.LanguageTagging, BookOnixAccessibilityFeature.LanguageTagging] },
            "feature" => new() { Features = [(BookOnixAccessibilityFeature)999] },
            "status" => new() { Status = (BookOnixAccessibilityStatus)999 },
            "version" => new() { Conformance = new((BookOnixWcagVersion)999, BookOnixWcagLevel.A) },
            "level" => new() { Conformance = new(BookOnixWcagVersion.V2_2, (BookOnixWcagLevel)999) },
            "url" => new() { PublisherInformationUrl = "javascript:alert(1)" },
            "credentials" => new() { PublisherInformationUrl = "https://name:secret@example.org/" },
            "email" => new() { PublisherContactEmail = "Publisher <accessibility@example.org>" },
            "oversized" => new() { Summary = new string('x', 4097) },
            _ => new()
        };
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Accessibility = metadata }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void CertificationAndReportsRetainTheirDistinctRolesWithoutInferringConformance() {
        var metadata = new BookOnixAccessibilityMetadata {
            Certification = new("Certifier & Partners", "https://certifier.example.org/scheme") {
                CredentiallingOrganizationName = "Credentials <Organization>",
                CredentiallingOrganizationUrl = "https://credentials.example.org/"
            },
            IndependentReportUrl = "https://certifier.example.org/report?a=1&b=2",
            IntermediaryInformationUrl = "https://intermediary.example.org/edition",
            PublisherInformationUrl = "https://publisher.example.org/edition",
            CompatibilityReportUrl = "https://publisher.example.org/compatibility",
            IntermediaryContactEmail = "support@intermediary.example.org",
            PublisherContactEmail = "support@publisher.example.org"
        };
        var xml = Export(metadata);
        Assert.Equal(new[] { "88", "89", "90", "93", "94", "95", "96", "97", "98", "99" },
            xml.Descendants(Onix + "ProductFormFeatureValue").Select(e => e.Value));
        foreach (var pair in new[] {
            ("88", metadata.Certification.CredentiallingOrganizationName), ("89", metadata.Certification.CredentiallingOrganizationUrl),
            ("90", metadata.Certification.Name), ("93", metadata.Certification.Url), ("94", metadata.IndependentReportUrl),
            ("95", metadata.IntermediaryInformationUrl), ("96", metadata.PublisherInformationUrl),
            ("97", metadata.CompatibilityReportUrl), ("98", metadata.IntermediaryContactEmail), ("99", metadata.PublisherContactEmail)
        }) Assert.Equal(pair.Item2, Description(xml, pair.Item1));
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        var result = project.ExportOnix(BookOnixTests.Options() with { Accessibility = metadata }, BookOnixTests.TestSchema());
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(project.Export().Bytes, result.Publication.Bytes);
    }

    [Fact]
    public void IndependentReportDoesNotRequireOrInventACertifier() {
        var xml = Export(new() { IndependentReportUrl = "https://testing.example.org/report" });
        Assert.Equal("94", xml.Descendants(Onix + "ProductFormFeatureValue").Single().Value);
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("https://user:secret@example.org/")]
    [InlineData("relative/report")]
    [InlineData("https://example.org/a b")]
    [InlineData("")]
    public void EveryProvenanceUrlUsesTheSameValidation(string invalid) {
        BookOnixAccessibilityMetadata[] declarations = [
            new() { Certification = new("Certifier", invalid) },
            new() { Certification = new("Certifier", "https://example.org/") { CredentiallingOrganizationUrl = invalid } },
            new() { IndependentReportUrl = invalid }, new() { IntermediaryInformationUrl = invalid },
            new() { CompatibilityReportUrl = invalid }, new() { PublisherInformationUrl = invalid }
        ];
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        foreach (var metadata in declarations)
            Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Accessibility = metadata }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void InvalidCertifierNamesRequiredUrlsAndIntermediaryContactsAreRejected() {
        BookOnixAccessibilityMetadata[] declarations = [
            new() { Certification = new(null!, "https://example.org/") },
            new() { Certification = new("Certifier", null!) },
            new() { Certification = new(" ", "https://example.org/") },
            new() { Certification = new(new string('x', 4097), "https://example.org/") },
            new() { IntermediaryContactEmail = "Support <support@example.org>" },
            new() { IntermediaryContactEmail = "not-an-address" }
        ];
        foreach (var metadata in declarations) Assert.ThrowsAny<ArgumentException>(() => Export(metadata));
        Assert.Throws<System.Xml.XmlException>(() => Export(new() {
            Certification = new("Certifier", "https://example.org/") { CredentiallingOrganizationName = "bad\u0001name" }
        }));
    }

    private static XDocument Export(BookOnixAccessibilityMetadata metadata) => XDocument.Load(new MemoryStream(
        BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Accessibility = metadata }, BookOnixTests.TestSchema()).Bytes));
    private static string Description(XDocument xml, string code) => xml.Descendants(Onix + "ProductFormFeature")
        .Single(e => e.Element(Onix + "ProductFormFeatureValue")!.Value == code).Element(Onix + "ProductFormFeatureDescription")!.Value;
}
