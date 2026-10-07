using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixCollateralTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixCollateralText Item() => new() { Type = BookOnixTextType.ReviewQuote,
        Audiences = [BookOnixContentAudience.EndCustomers, BookOnixContentAudience.Librarians],
        Texts = [new("A <strong>literal</strong> & supplied quote", "eng"), new("Przykładowa opinia", "pol")],
        Authors = ["Example Reviewer"], SourceCorporate = "Example Reviews",
        SourceTitles = [new("Review & journal", "eng")], SourceLinks = ["https://example.org/review?a=1&b=2"],
        Territory = new() { Countries = ["US", "PL"] }, PublishedOn = new(2026, 9, 1),
        UsableFrom = new(2026, 10, 1), UsableUntil = new(2027, 10, 1), UpdatedOn = new(2026, 9, 2) };

    [Fact]
    public void TextAttributionRecipientTerritoryAndDatesArePreservedWithoutChangingTheBook() {
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item(), Item() with {
            Type = BookOnixTextType.Description, Texts = [new(new string('A', 8192))] }] }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes)); var collateral = xml.Descendants(Ns + "CollateralDetail").Single();
        var entries = collateral.Elements(Ns + "TextContent").ToArray();
        Assert.Equal(new[] { "1", "2" }, entries.Select(e => e.Element(Ns + "SequenceNumber")!.Value));
        Assert.Equal(new[] { "06", "03" }, entries.Select(e => e.Element(Ns + "TextType")!.Value));
        Assert.Equal(new[] { "03", "04" }, entries[0].Elements(Ns + "ContentAudience").Select(e => e.Value));
        Assert.Equal("PL US", entries[0].Descendants(Ns + "CountriesIncluded").Single().Value);
        var texts = entries[0].Elements(Ns + "Text").ToArray();
        Assert.Equal(new[] { "eng", "pol" }, texts.Select(e => (string?)e.Attribute("language")));
        Assert.All(texts, e => { Assert.Equal("06", (string?)e.Attribute("textformat")); Assert.Empty(e.Elements()); });
        Assert.Equal(Item().Texts[0].Text, texts[0].Value);
        Assert.Equal("Example Reviewer", entries[0].Element(Ns + "TextAuthor")!.Value);
        Assert.Equal("Example Reviews", entries[0].Element(Ns + "TextSourceCorporate")!.Value);
        Assert.Equal("eng", (string?)entries[0].Element(Ns + "SourceTitle")!.Attribute("language"));
        Assert.Equal(Item().SourceLinks[0], entries[0].Element(Ns + "TextSourceLink")!.Value);
        Assert.Equal(new[] { "01", "14", "15", "17" }, entries[0].Descendants(Ns + "ContentDateRole").Select(e => e.Value));
        Assert.Equal(new[] { "20260901", "20261001", "20271001", "20260902" }, entries[0].Descendants(Ns + "Date").Select(e => e.Value));
        Assert.Equal("PublishingDetail", collateral.ElementsAfterSelf().First().Name.LocalName);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData(BookOnixTextType.ShortDescription)]
    [InlineData(BookOnixTextType.CollectionShortDescription)]
    public void ShortDescriptionLimitCountsUnicodeScalarsRatherThanUtf16Units(BookOnixTextType type) {
        string text = string.Concat(Enumerable.Repeat("😀", 350));
        var project = BookOnixTests.Project();
        var options = BookOnixTests.Options() with { CollateralTexts = [Item() with { Type = type, Texts = [new(text)] }] };
        var result = project.ExportOnix(options, BookOnixTests.TestSchema());
        Assert.Equal(text, XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Text").Single().Value);
        Assert.Throws<ArgumentException>(() => project.ExportOnix(options with { CollateralTexts = [Item() with {
            Type = type, Texts = [new(text + "A")] }] }, BookOnixTests.TestSchema()));
    }

    [Theory]
    [InlineData("type")]
    [InlineData("no-audience")]
    [InlineData("audience")]
    [InlineData("duplicate-audience")]
    [InlineData("unrestricted-mixed")]
    [InlineData("no-text")]
    [InlineData("blank")]
    [InlineData("text-length")]
    [InlineData("text-count")]
    [InlineData("language")]
    [InlineData("duplicate-language")]
    [InlineData("authors")]
    [InlineData("source-language")]
    [InlineData("link-scheme")]
    [InlineData("link-credentials")]
    [InlineData("duplicate-links")]
    [InlineData("dates")]
    [InlineData("territory")]
    [InlineData("count")]
    [InlineData("budget")]
    public void InvalidCollateralFailsWithoutProjectMutation(string kind) {
        var item = kind switch {
            "type" => Item() with { Type = (BookOnixTextType)99 },
            "no-audience" => Item() with { Audiences = [] },
            "audience" => Item() with { Audiences = [(BookOnixContentAudience)99] },
            "duplicate-audience" => Item() with { Audiences = [BookOnixContentAudience.EndCustomers, BookOnixContentAudience.EndCustomers] },
            "unrestricted-mixed" => Item() with { Audiences = [BookOnixContentAudience.Unrestricted, BookOnixContentAudience.EndCustomers] },
            "no-text" => Item() with { Texts = [] },
            "blank" => Item() with { Texts = [new(" ")] },
            "text-length" => Item() with { Texts = [new(new string('A', 65537))] },
            "text-count" => Item() with { Texts = Enumerable.Repeat(new BookOnixCollateralTextValue("A"), 17).ToArray() },
            "language" => Item() with { Texts = [new("A", "en-US")] },
            "duplicate-language" => Item() with { Texts = [new("A"), new("B")] },
            "authors" => Item() with { Authors = ["Name", "Name"] },
            "source-language" => Item() with { SourceTitles = [new("A", "eng"), new("B", "eng")] },
            "link-scheme" => Item() with { SourceLinks = ["file:///private/review"] },
            "link-credentials" => Item() with { SourceLinks = ["https://user:password@example.org/review"] },
            "duplicate-links" => Item() with { SourceLinks = ["https://example.org/review", "https://example.org/review"] },
            "dates" => Item() with { UsableUntil = new(2026, 1, 1) },
            "territory" => Item() with { Territory = new() },
            "budget" => Item() with { Texts = [new(new string('A', 65536))], SourceTitles = [] },
            _ => Item()
        };
        var items = kind is "count" or "budget" ? Enumerable.Repeat(item, kind == "count" ? 65 : 9).ToArray() : [item];
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = items }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}
