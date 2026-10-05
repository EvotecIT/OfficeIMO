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

    private static XDocument Export(BookOnixAccessibilityMetadata metadata) => XDocument.Load(new MemoryStream(
        BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Accessibility = metadata }, BookOnixTests.TestSchema()).Bytes));
    private static string Description(XDocument xml, string code) => xml.Descendants(Onix + "ProductFormFeature")
        .Single(e => e.Element(Onix + "ProductFormFeatureValue")!.Value == code).Element(Onix + "ProductFormFeatureDescription")!.Value;
}
