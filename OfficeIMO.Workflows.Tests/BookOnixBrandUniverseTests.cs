using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixBrandUniverseTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixCollectionTitleElement Title(BookOnixCollectionLevel level, string text) => new() { Level = level, Title = text };

    [Theory]
    [InlineData(BookOnixCollectionLevel.MasterBrand, "05")]
    [InlineData(BookOnixCollectionLevel.Universe, "07")]
    public void NamedGroupingDoesNotRequireAnInventedSeries(BookOnixCollectionLevel level, string code) {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Collections = [new() {
            Type = BookOnixCollectionType.Publisher,
            TitleElements = [Title(level, "The Star Garden") with { LanguageCode = "eng", Subtitle = "A shared identity", TitleSorting = new() { Prefix = "The " } }],
            Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "star-garden", "Publisher catalog") { Level = level }]
        }] }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var collection = xml.Descendants(Ns + "Collection").Single();
        Assert.Equal(code, collection.Descendants(Ns + "TitleElementLevel").Single().Value);
        Assert.Equal(code, collection.Descendants(Ns + "CollectionElementLevel").Single().Value);
        Assert.Equal("The ", collection.Descendants(Ns + "TitlePrefix").Single().Value);
        Assert.Equal("Star Garden", collection.Descendants(Ns + "TitleWithoutPrefix").Single().Value);
        Assert.Equal("eng", (string?)collection.Descendants(Ns + "Subtitle").Single().Attribute("language"));
        Assert.Equal("star-garden", collection.Descendants(Ns + "IDValue").Single().Value);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Fact]
    public void BrandingAndSeriesLevelsKeepExplicitDisplayOrderAndIdentifierScopes() {
        var levels = new[] { BookOnixCollectionLevel.Universe, BookOnixCollectionLevel.Subcollection,
            BookOnixCollectionLevel.MasterBrand, BookOnixCollectionLevel.Collection, BookOnixCollectionLevel.SubSubcollection };
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Collections = [new() {
            Type = BookOnixCollectionType.Publisher,
            TitleElements = levels.Select(level => Title(level, level.ToString())).ToArray(),
            Identifiers = levels.Select(level => new BookOnixCollectionIdentifier(BookOnixCollectionIdentifierType.Proprietary,
                level.ToString(), "Catalog") { Level = level }).ToArray()
        }] }, BookOnixTests.TestSchema());
        var collection = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Collection").Single();
        Assert.Equal(new[] { "07", "03", "05", "02", "06" }, collection.Descendants(Ns + "TitleElementLevel").Select(e => e.Value));
        Assert.Equal(new[] { "07", "03", "05", "02", "06" }, collection.Descendants(Ns + "CollectionElementLevel").Select(e => e.Value));
        Assert.Equal(new[] { "1", "2", "3", "4", "5" }, collection.Descendants(Ns + "SequenceNumber").Select(e => e.Value));
    }

    [Fact]
    public void IndependentBrandsRemainSeparateMemberships() {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Collections = new[] { "Brand A", "Brand B" }
            .Select(name => new BookOnixCollection { Type = BookOnixCollectionType.Publisher,
                TitleElements = [Title(BookOnixCollectionLevel.MasterBrand, name)] }).ToArray()
        }, BookOnixTests.TestSchema());
        var collections = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Collection").ToArray();
        Assert.Equal(2, collections.Length);
        Assert.Equal(new[] { "Brand A", "Brand B" }, collections.Select(c => c.Descendants(Ns + "TitleText").Single().Value));
    }

    [Theory]
    [InlineData("brand-without-name")]
    [InlineData("universe-without-name")]
    [InlineData("duplicate-brand")]
    [InlineData("missing-series-root")]
    [InlineData("missing-series-parent")]
    [InlineData("missing-identifier-level")]
    [InlineData("no-collection-conflict")]
    public void BrandingCannotBypassNameHierarchyAndIdentityRules(string invalid) {
        var brand = Title(BookOnixCollectionLevel.MasterBrand, "Brand");
        var collection = new BookOnixCollection { Type = BookOnixCollectionType.Publisher, TitleElements = [brand] };
        collection = invalid switch {
            "brand-without-name" => collection with { TitleElements = [brand with { Title = null, PartNumber = "1" }] },
            "universe-without-name" => collection with { TitleElements = [brand with { Level = BookOnixCollectionLevel.Universe, Title = null, PartNumber = "1" }] },
            "duplicate-brand" => collection with { TitleElements = [brand, brand] },
            "missing-series-root" => collection with { TitleElements = [brand, Title(BookOnixCollectionLevel.Subcollection, "Subseries")] },
            "missing-series-parent" => collection with { TitleElements = [brand, Title(BookOnixCollectionLevel.Collection, "Series"), Title(BookOnixCollectionLevel.SubSubcollection, "Part")] },
            "missing-identifier-level" => collection with { Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "id", "Catalog") { Level = BookOnixCollectionLevel.Universe }] },
            _ => collection
        };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with {
            Collections = [collection], NoCollection = invalid == "no-collection-conflict"
        }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}
