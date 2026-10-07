using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixAudienceHeadingTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;

    [Fact]
    public void RepeatedDescriptionsRequireLanguagesWithoutMutation() {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        Assert.Throws<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with {
            Audience = new() { Descriptions = [new("Readers", "eng"), new("Czytelnicy")] }
        }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void CategoryAndSchemeHeadingsPreserveTranslationsAndPublicationBytes() {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Categories = [new(BookOnixAudienceType.Children, true) { Headings = [new("Children & families", "eng"), new("Dzieci", "pol")] }],
            Codes = [new(BookOnixAudienceScheme.Proprietary, "family") { SchemeName = "House", IsMain = true,
                Headings = [new(" Families ", "eng"), new("Rodziny", "pol")] }]
        } }, BookOnixTests.TestSchema());
        var audiences = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Audience").ToArray();
        Assert.Equal(new[] { "02", "family" }, audiences.Select(e => e.Element(Ns + "AudienceCodeValue")!.Value));
        Assert.Equal(new[] { "Children & families", "Dzieci", " Families ", "Rodziny" }, audiences.SelectMany(e => e.Elements(Ns + "AudienceHeadingText")).Select(e => e.Value));
        Assert.All(audiences, element => {
            Assert.Equal(new[] { "eng", "pol" }, element.Elements(Ns + "AudienceHeadingText").Select(e => (string?)e.Attribute("language")));
            Assert.Equal("AudienceHeadingText", element.Elements().Last().Name.LocalName);
        });
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Fact]
    public void DistinctHeadingOnlyAssertionsDoNotInventCodeValues() {
        var result = Export(new() { Codes = [
            new(BookOnixAudienceScheme.Proprietary) { SchemeName = "House", Headings = [new("Families")] },
            new(BookOnixAudienceScheme.Proprietary) { SchemeName = "House", Headings = [new("Teachers")] },
            new(BookOnixAudienceScheme.Proprietary) { SchemeName = "Recipient", Headings = [new("Families")] }
        ], Descriptions = [new("Reader groups")] });
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        Assert.Empty(xml.Descendants(Ns + "AudienceCodeValue"));
        Assert.Equal(new[] { "Families", "Teachers", "Families" }, xml.Descendants(Ns + "AudienceHeadingText").Select(e => e.Value));
        Assert.All(xml.Descendants(Ns + "AudienceHeadingText"), e => Assert.Null(e.Attribute("language")));
        Assert.Null(xml.Descendants(Ns + "AudienceDescription").Single().Attribute("language"));
    }

    [Theory]
    [InlineData(BookOnixAudienceScheme.Cefr)]
    [InlineData(BookOnixAudienceScheme.JapaneseChildren)]
    [InlineData(BookOnixAudienceScheme.IntendedLanguage)]
    public void HeadingOnlyAssertionsDoNotInferOrValidateAnAbsentCode(BookOnixAudienceScheme scheme) {
        var result = Export(new() { Codes = [new(scheme) { Headings = [new("Publisher-supplied equivalent", "eng")] }] });
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        Assert.Empty(xml.Descendants(Ns + "AudienceCodeValue"));
        Assert.Single(xml.Descendants(Ns + "AudienceHeadingText"));
    }

    [Theory]
    [InlineData("missing-language")]
    [InlineData("duplicate-language")]
    [InlineData("language")]
    [InlineData("empty")]
    [InlineData("oversized")]
    [InlineData("count")]
    [InlineData("null-list")]
    [InlineData("null-heading")]
    [InlineData("category-missing-language")]
    [InlineData("category-null-list")]
    public void InvalidHeadingTranslationsFailWithoutMutation(string kind) {
        IReadOnlyList<BookOnixAudienceHeading> headings = kind switch {
            "missing-language" or "category-missing-language" => [new("One", "eng"), new("Dwa")],
            "duplicate-language" => [new("One", "eng"), new("Two", "eng")],
            "language" => [new("One", "en")],
            "empty" => [new(" ")],
            "oversized" => [new(new string('x', 4097))],
            "count" => Enumerable.Repeat(new BookOnixAudienceHeading("One", "eng"), 17).ToArray(),
            "null-list" or "category-null-list" => null!,
            _ => [null!]
        };
        var audience = kind.StartsWith("category", StringComparison.Ordinal)
            ? new BookOnixAudienceMetadata { Categories = [new(BookOnixAudienceType.Children) { Headings = headings }] }
            : new BookOnixAudienceMetadata { Codes = [new(BookOnixAudienceScheme.Proprietary, "one") { SchemeName = "House", Headings = headings }] };
        Reject(audience);
    }

    [Theory]
    [InlineData("no-content")]
    [InlineData("empty-code")]
    [InlineData("duplicate-heading")]
    [InlineData("duplicate-code")]
    [InlineData("multiple-main")]
    public void UncodedAssertionsRetainIdentityAndMainAudienceRules(string kind) {
        var first = new BookOnixAudienceCode(BookOnixAudienceScheme.Proprietary) {
            SchemeName = "House", Headings = [new("Families", "eng"), new("Rodziny", "pol")] };
        IReadOnlyList<BookOnixAudienceCode> codes = kind switch {
            "no-content" => [first with { Headings = [] }],
            "empty-code" => [first with { Value = "" }],
            "duplicate-heading" => [first, first with { Headings = first.Headings.Reverse().ToArray() }],
            "duplicate-code" => [first with { Value = "one" }, first with { Value = "one", Headings = [new("Other")] }],
            _ => [first with { IsMain = true }, first with { IsMain = true, Value = "one" }]
        };
        Reject(new() { Codes = codes });
    }

    private static BookOnixExportResult Export(BookOnixAudienceMetadata audience) =>
        BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Audience = audience }, BookOnixTests.TestSchema());

    private static void Reject(BookOnixAudienceMetadata audience) {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Audience = audience }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}
