using System.Xml.Linq;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixSubterritoryTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixTerritory Country => new() { Countries = ["US"] };
    private static BookOnixTerritory California => new() { Regions = ["US-CA"] };
    private static BookOnixCommercialMetadata Commercial(BookOnixTerritory market, params BookOnixSalesRights[] rights) => new() {
        PublishingStatus = BookOnixPublishingStatus.Active, SalesRights = rights,
        Supplies = [new() { Territory = market, SupplierName = "Example Supplier",
            SupplierRole = BookOnixSupplierRole.PublisherToResellers, Availability = BookOnixAvailability.Available,
            Prices = [new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, Amount = 10, CurrencyCode = "USD" }] }]
    };
    private static XDocument Export(BookOnixCommercialMetadata commercial) => XDocument.Load(new MemoryStream(
        BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Commercial = commercial }, BookOnixTests.TestSchema()).Bytes));

    [Fact]
    public void CountryRightsCoverRegionalMarketAndInheritedPrice() {
        var xml = Export(Commercial(California, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, Country)));
        Assert.Equal("US-CA", xml.Descendants(Ns + "Market").Single().Descendants(Ns + "RegionsIncluded").Single().Value);
        Assert.Equal("US-CA", xml.Descendants(Ns + "Price").Single().Descendants(Ns + "RegionsIncluded").Single().Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DisjointRegionalGrantsCoverCountryOrWorldWithoutOverlap(bool world) {
        var whole = world ? new BookOnixTerritory { Worldwide = true } : Country;
        var xml = Export(Commercial(whole,
            new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, whole with { ExcludedRegions = ["US-CA"] }),
            new BookOnixSalesRights(BookOnixSalesRightsKind.NonExclusive, California)));
        Assert.Equal("US-CA", xml.Descendants(Ns + "RegionsExcluded").Single().Value);
    }

    [Fact]
    public void MixedTerritorySerializesInSchemaOrderWithSortedCodes() {
        var territory = new BookOnixTerritory { Countries = ["US", "GB"], Regions = ["CA-QC", "CA-ON"], ExcludedRegions = ["US-NY", "US-CA"] };
        var xml = Export(Commercial(territory, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, territory)));
        var node = xml.Descendants(Ns + "Territory").First();
        Assert.Equal(new[] { "CountriesIncluded", "RegionsIncluded", "RegionsExcluded" }, node.Elements().Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "GB US", "CA-ON CA-QC", "US-CA US-NY" }, node.Elements().Select(e => e.Value));
    }

    [Theory]
    [InlineData("regional-grant")]
    [InlineData("world-grant")]
    [InlineData("excluded")]
    [InlineData("denied")]
    [InlineData("overlap")]
    [InlineData("price")]
    public void IncompleteOrConflictingRightsNeverBroadenMarket(string kind) {
        var commercial = kind switch {
            "regional-grant" => Commercial(Country, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, California)),
            "world-grant" => Commercial(new() { Worldwide = true }, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, California)),
            "excluded" => Commercial(California, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, Country with { ExcludedRegions = ["US-CA"] })),
            "denied" => Commercial(California, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, Country with { ExcludedRegions = ["US-CA"] }), new BookOnixSalesRights(BookOnixSalesRightsKind.RightsNotHeld, California)),
            "overlap" => Commercial(California, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, Country), new BookOnixSalesRights(BookOnixSalesRightsKind.NonExclusive, California)),
            _ => Commercial(California, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, Country))
        };
        if (kind == "price") commercial = commercial with { Supplies = [commercial.Supplies[0] with {
            Prices = [commercial.Supplies[0].Prices[0] with { Territory = new() { Regions = ["US-NY"] } }]
        }] };
        Reject(commercial);
    }

    [Theory]
    [InlineData("WORLD")]
    [InlineData("ECZ")]
    [InlineData("GB-EWS")]
    [InlineData("GB-LHR")]
    [InlineData("US-ZZ")]
    [InlineData("us-ca")]
    [InlineData("CN-11")]
    [InlineData("CN-HK")]
    [InlineData("")]
    [InlineData(null)]
    public void UnsupportedAliasesAndUnknownRegionsAreRejected(string? code) =>
        Reject(Commercial(California, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, new() { Regions = [code!] })));

    [Theory]
    [InlineData("duplicate")]
    [InlineData("parent")]
    [InlineData("world")]
    [InlineData("outside")]
    [InlineData("excluded-parent")]
    [InlineData("null")]
    [InlineData("limit")]
    public void InvalidRegionalConfigurationsAreRejected(string kind) {
        var territory = kind switch {
            "duplicate" => California with { Regions = ["US-CA", "US-CA"] },
            "parent" => Country with { Regions = ["US-CA"] },
            "world" => California with { Worldwide = true },
            "outside" => Country with { ExcludedRegions = ["CA-QC"] },
            "excluded-parent" => new BookOnixTerritory { Worldwide = true, ExcludedCountries = ["US"], ExcludedRegions = ["US-CA"] },
            "null" => Country with { ExcludedRegions = null! },
            _ => Country with { ExcludedRegions = Enumerable.Repeat("US-CA", 251).ToArray() }
        };
        Reject(Commercial(California, new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, territory)));
    }

    private static void Reject(BookOnixCommercialMetadata commercial) {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Commercial = commercial }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}
