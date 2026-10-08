using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixCollectionIdentifierTests {
    [Theory]
    [InlineData(BookOnixCollectionIdentifierType.GermanNationalBibliography, "series-001", "03", "series-001")]
    [InlineData(BookOnixCollectionIdentifierType.GermanBooksInPrint, "001234", "04", "001234")]
    [InlineData(BookOnixCollectionIdentifierType.Electre, "collection-é", "05", "collection-é")]
    [InlineData(BookOnixCollectionIdentifierType.Doi, "10.1000.12/Series α: Part 1", "06", "10.1000.12/Series α: Part 1")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:example:series%20one", "22", "urn:example:series%20one")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "URN:example:series/one?+resolve?=query#part", "22", "URN:example:series/one?+resolve?=query#part")]
    [InlineData(BookOnixCollectionIdentifierType.JapaneseMagazine, "01234", "27", "01234")]
    [InlineData(BookOnixCollectionIdentifierType.BnfControlNumber, "FRBNF-001", "29", "FRBNF-001")]
    [InlineData(BookOnixCollectionIdentifierType.Ark, "https://example.org/catalog/ark:/12345/series-1", "35", "https://example.org/catalog/ark:/12345/series-1")]
    [InlineData(BookOnixCollectionIdentifierType.IssnL, "2434-561x", "38", "2434561X")]
    public void StandardSchemesPreserveExplicitIdentityAndScope(BookOnixCollectionIdentifierType type, string value, string code, string expected) {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(Options(new BookOnixCollectionIdentifier(type, value) { Level = BookOnixCollectionLevel.Collection }), BookOnixTests.TestSchema());
        XNamespace ns = BookProject.OnixNamespace;
        var identifier = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(ns + "CollectionIdentifier").Single();
        Assert.Equal(code, identifier.Element(ns + "CollectionIDType")!.Value);
        Assert.Equal(expected, identifier.Element(ns + "IDValue")!.Value);
        Assert.Equal("02", identifier.Element(ns + "CollectionElementLevel")!.Value);
        Assert.Null(identifier.Element(ns + "IDTypeName"));
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Theory]
    [InlineData(BookOnixCollectionIdentifierType.Doi, "https://doi.org/10.1000/series")]
    [InlineData(BookOnixCollectionIdentifierType.Doi, "10.1000/")]
    [InlineData(BookOnixCollectionIdentifierType.Doi, "10..1000/a")]
    [InlineData(BookOnixCollectionIdentifierType.Doi, "10.abc/a")]
    [InlineData(BookOnixCollectionIdentifierType.Doi, "10.1000/a\nb")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:example:")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:x:value")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:example:/value")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:example:bad%2")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:example:with space")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:example:value?query")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:example:value?+resolve?=")]
    [InlineData(BookOnixCollectionIdentifierType.Urn, "urn:example:value?+resolve?=/query")]
    [InlineData(BookOnixCollectionIdentifierType.JapaneseMagazine, "12345-01")]
    [InlineData(BookOnixCollectionIdentifierType.JapaneseMagazine, "１２３４５")]
    [InlineData(BookOnixCollectionIdentifierType.Ark, "ark:/12345/series")]
    [InlineData(BookOnixCollectionIdentifierType.Ark, "https://example.org/series")]
    [InlineData(BookOnixCollectionIdentifierType.Ark, "https://example.org?redirect=/ark:/12345/series")]
    [InlineData(BookOnixCollectionIdentifierType.Ark, "https://example.org/ark:/abc/series")]
    [InlineData(BookOnixCollectionIdentifierType.Ark, "https://example.org/ark:/abc/../12345/series")]
    [InlineData(BookOnixCollectionIdentifierType.Ark, "https://example.org/ark:/12345/")]
    [InlineData(BookOnixCollectionIdentifierType.IssnL, "2434-5610")]
    public void MalformedIdentifiersFailWithoutMutatingTheProject(BookOnixCollectionIdentifierType type, string value) {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(Options(new BookOnixCollectionIdentifier(type, value)), BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void StandardSchemesCannotAcquireProprietaryNamesOrOverlapScopes() {
        var project = BookOnixTests.Project();
        Assert.Throws<ArgumentException>(() => project.ExportOnix(Options(new BookOnixCollectionIdentifier(BookOnixCollectionIdentifierType.Electre, "one", "custom")), BookOnixTests.TestSchema()));
        var identifier = new BookOnixCollectionIdentifier(BookOnixCollectionIdentifierType.Doi, "10.1000/one");
        Assert.Throws<ArgumentException>(() => project.ExportOnix(Options(identifier, identifier with { Level = BookOnixCollectionLevel.Collection }), BookOnixTests.TestSchema()));
    }

    private static BookOnixExportOptions Options(params BookOnixCollectionIdentifier[] identifiers) => BookOnixTests.Options() with {
        Collections = [new() { Type = BookOnixCollectionType.Publisher, Title = "Series", Identifiers = identifiers }]
    };
}
