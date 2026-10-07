using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixCollectionTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixCollection Collection() => new() { Type = BookOnixCollectionType.Publisher, Title = "Series & studies",
        Subtitle = "A collection", LanguageCode = "eng",
        Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "series-1", "Publisher catalog"),
            new(BookOnixCollectionIdentifierType.Issn, "1092-003x")],
        Sequences = [new(BookOnixCollectionSequenceType.Narrative, "2.1"),
            new(BookOnixCollectionSequenceType.Proprietary, "3.-.8", "Curriculum order")] };

    [Fact]
    public void CollectionsPreserveTitleIdentityAndMultipleSequenceMeanings() {
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Collections = [Collection(),
            new() { Type = BookOnixCollectionType.Ascribed, Title = "Curated", SourceName = "Library",
                Identifiers = [new(BookOnixCollectionIdentifierType.Isbn13, "978-0-306-40615-7")] }] }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var collections = xml.Descendants(Ns + "Collection").ToArray();
        Assert.Equal(new[] { "10", "20" }, collections.Select(e => e.Element(Ns + "CollectionType")!.Value));
        Assert.Equal(new[] { "series-1", "1092003X" }, collections[0].Descendants(Ns + "IDValue").Select(e => e.Value));
        Assert.Equal("9780306406157", collections[1].Descendants(Ns + "IDValue").Single().Value);
        Assert.Equal(new[] { "04", "01" }, collections[0].Descendants(Ns + "CollectionSequenceType").Select(e => e.Value));
        Assert.Equal(new[] { "2.1", "3.-.8" }, collections[0].Descendants(Ns + "CollectionSequenceNumber").Select(e => e.Value));
        Assert.Equal("02", collections[0].Descendants(Ns + "TitleElementLevel").Single().Value);
        Assert.Null(collections[0].Descendants(Ns + "TitleElement").Single().Attribute("language"));
        Assert.Equal("eng", (string?)collections[0].Descendants(Ns + "TitleText").Single().Attribute("language"));
        Assert.Equal("eng", (string?)collections[0].Descendants(Ns + "Subtitle").Single().Attribute("language"));
        Assert.Equal("Series & studies", collections[0].Descendants(Ns + "TitleText").Single().Value);
        Assert.Equal("Book & title", xml.Descendants(Ns + "DescriptiveDetail").Single().Elements(Ns + "TitleDetail").Single().Descendants(Ns + "TitleText").Single().Value);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData("0317-8471", "03178471")]
    [InlineData("2434-561x", "2434561X")]
    [InlineData("0378-5955", "03785955")]
    public void IssnChecksumAndOptionalHyphenAreNormalized(string input, string expected) {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Collections = [Collection() with {
            Identifiers = [new(BookOnixCollectionIdentifierType.Issn, input)] }] }, BookOnixTests.TestSchema());
        Assert.Equal(expected, XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "CollectionIdentifier").Single().Element(Ns + "IDValue")!.Value);
    }

    [Fact]
    public void NoCollectionIsAnExplicitAssertion() {
        var project = BookOnixTests.Project();
        Assert.Empty(XDocument.Load(new MemoryStream(project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema()).Bytes)).Descendants(Ns + "NoCollection"));
        Assert.Single(XDocument.Load(new MemoryStream(project.ExportOnix(BookOnixTests.Options() with { NoCollection = true }, BookOnixTests.TestSchema()).Bytes)).Descendants(Ns + "NoCollection"));
        Assert.Throws<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { NoCollection = true, Collections = [Collection()] }, BookOnixTests.TestSchema()));
    }

    [Theory]
    [InlineData("type")]
    [InlineData("title")]
    [InlineData("language")]
    [InlineData("source")]
    [InlineData("identifier-type")]
    [InlineData("identifier-name")]
    [InlineData("standard-name")]
    [InlineData("issn-checksum")]
    [InlineData("issn-shape")]
    [InlineData("isbn-checksum")]
    [InlineData("duplicate-identifier")]
    [InlineData("sequence-type")]
    [InlineData("sequence-name")]
    [InlineData("sequence-standard-name")]
    [InlineData("sequence-shape")]
    [InlineData("duplicate-sequence")]
    [InlineData("identifier-count")]
    [InlineData("sequence-count")]
    [InlineData("collection-count")]
    public void InvalidCollectionAssertionsFailWithoutMutation(string kind) {
        var collection = kind switch {
            "type" => Collection() with { Type = (BookOnixCollectionType)99 },
            "title" => Collection() with { Title = " " },
            "language" => Collection() with { LanguageCode = "en-US" },
            "source" => Collection() with { Type = BookOnixCollectionType.Ascribed },
            "identifier-type" => Collection() with { Identifiers = [new((BookOnixCollectionIdentifierType)99, "id")] },
            "identifier-name" => Collection() with { Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "id")] },
            "standard-name" => Collection() with { Identifiers = [new(BookOnixCollectionIdentifierType.Issn, "03178471", "Custom")] },
            "issn-checksum" => Collection() with { Identifiers = [new(BookOnixCollectionIdentifierType.Issn, "03178472")] },
            "issn-shape" => Collection() with { Identifiers = [new(BookOnixCollectionIdentifierType.Issn, "0317847:")] },
            "isbn-checksum" => Collection() with { Identifiers = [new(BookOnixCollectionIdentifierType.Isbn13, "9780306406158")] },
            "duplicate-identifier" => Collection() with { Identifiers = [new(BookOnixCollectionIdentifierType.Issn, "03178471"), new(BookOnixCollectionIdentifierType.Issn, "1092003X")] },
            "sequence-type" => Collection() with { Sequences = [new((BookOnixCollectionSequenceType)99, "1")] },
            "sequence-name" => Collection() with { Sequences = [new(BookOnixCollectionSequenceType.Proprietary, "1")] },
            "sequence-standard-name" => Collection() with { Sequences = [new(BookOnixCollectionSequenceType.Narrative, "1", "Named")] },
            "sequence-shape" => Collection() with { Sequences = [new(BookOnixCollectionSequenceType.Narrative, "2..1")] },
            "duplicate-sequence" => Collection() with { Sequences = [new(BookOnixCollectionSequenceType.Narrative, "1"), new(BookOnixCollectionSequenceType.Narrative, "2")] },
            "identifier-count" => Collection() with { Identifiers = Enumerable.Repeat(new BookOnixCollectionIdentifier(BookOnixCollectionIdentifierType.Proprietary, "id", "Name"), 17).ToArray() },
            "sequence-count" => Collection() with { Sequences = Enumerable.Repeat(new BookOnixCollectionSequence(BookOnixCollectionSequenceType.Narrative, "1"), 17).ToArray() },
            _ => Collection()
        };
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        var collections = kind == "collection-count" ? Enumerable.Repeat(collection, 33).ToArray() : [collection];
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Collections = collections }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}
