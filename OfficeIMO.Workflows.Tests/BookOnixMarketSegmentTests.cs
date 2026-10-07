using System.Xml.Linq;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixMarketSegmentTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixTerritory Us => new() { Countries = ["US"] };
    private static BookOnixSupply Supply() => new() {
        Territory = Us, MarketReference = "us-offer", SupplierName = "Example Supplier",
        SupplierRole = BookOnixSupplierRole.PublisherToResellers, Availability = BookOnixAvailability.Available,
        Prices = [new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, CurrencyCode = "USD", Amount = 12 }],
        MarketSegments = [
            new() { Territory = new() { Regions = ["US-CA"] }, Restrictions = [new(BookOnixSalesRestrictionKind.LibrariesOnly)] },
            new() { Territory = Us with { ExcludedRegions = ["US-CA"] }, Restrictions = [new(BookOnixSalesRestrictionKind.ExceptLibraries)] }
        ]
    };
    private static BookOnixExportResult Export(BookOnixSupply supply, BookProject? project = null) =>
        (project ?? BookOnixTests.Project()).ExportOnix(Options(supply), BookOnixTests.TestSchema(),
            new OfficeIMO.Epub.EpubWriteOptions { ModifiedAt = BookOnixTests.Options().SentAt });
    private static BookOnixExportOptions Options(BookOnixSupply supply) => BookOnixTests.Options() with {
        Commercial = new() { PublishingStatus = BookOnixPublishingStatus.Active,
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, Us)], Supplies = [supply] }
    };

    [Fact]
    public void DisjointChannelsShareOneSupplyPriceAndStableIdentity() {
        var project = BookOnixTests.Project();
        var result = Export(Supply(), project);
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var supply = xml.Descendants(Ns + "ProductSupply").Single();
        Assert.Equal("us-offer", supply.Element(Ns + "MarketReference")!.Value);
        Assert.Equal(new[] { "06", "09" }, supply.Elements(Ns + "Market").Select(m => m.Descendants(Ns + "SalesRestrictionType").Single().Value));
        Assert.Single(supply.Elements(Ns + "SupplyDetail"));
        Assert.Equal("US", supply.Descendants(Ns + "Price").Single().Descendants(Ns + "CountriesIncluded").Single().Value);
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
        var update = BookOnixMessage.CreateBlockUpdates([new(result) { ReplaceMarketReferences = ["us-offer"] }], BookOnixTests.TestSchema());
        var updateXml = XDocument.Load(new MemoryStream(update.Bytes));
        Assert.True(XNode.DeepEquals(supply, updateXml.Descendants(Ns + "ProductSupply").Single()));
        Assert.Equal(Export(Supply() with { MarketSegments = [] }, project).Publication.Bytes, result.Publication.Bytes);
    }

    [Fact]
    public void SupplyWideRestrictionsAreRetainedInEverySegment() {
        var xml = XDocument.Load(new MemoryStream(Export(Supply() with {
            Restrictions = [new(BookOnixSalesRestrictionKind.ExceptOnlineRetail)]
        }).Bytes));
        Assert.All(xml.Descendants(Ns + "Market"), market => Assert.Equal("14", market.Descendants(Ns + "SalesRestrictionType").First().Value));
    }

    [Theory]
    [InlineData("gap")]
    [InlineData("overlap")]
    [InlineData("outside")]
    [InlineData("opposed")]
    [InlineData("null")]
    [InlineData("count")]
    [InlineData("combined-count")]
    public void InvalidPartitionsAndCombinedRestrictionsLeaveProjectUnchanged(string kind) {
        var supply = Supply();
        supply = kind switch {
            "gap" => supply with { MarketSegments = [supply.MarketSegments[0]] },
            "overlap" => supply with { MarketSegments = [supply.MarketSegments[0], new() { Territory = Us }] },
            "outside" => supply with { MarketSegments = [new() { Territory = new() { Countries = ["CA"] } }] },
            "opposed" => supply with { Restrictions = [new(BookOnixSalesRestrictionKind.NoRestrictions)] },
            "null" => supply with { MarketSegments = null! },
            "count" => supply with { MarketSegments = Enumerable.Repeat(supply.MarketSegments[0], 33).ToArray() },
            _ => supply with { Restrictions = Enumerable.Repeat(new BookOnixSalesRestriction(BookOnixSalesRestrictionKind.LibrariesOnly), 32).ToArray() }
        };
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(Options(supply), BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}
