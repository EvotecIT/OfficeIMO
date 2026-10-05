using System.Globalization;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixCommercialTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixTerritory Countries(params string[] codes) => new() { Countries = codes };
    private static BookOnixPrice Price() => new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, Amount = 12.3400m, CurrencyCode = "PLN" };
    private static BookOnixSupply Supply() => new() {
        Territory = Countries("PL"), SupplierName = "Example Supplier", SupplierRole = BookOnixSupplierRole.PublisherToResellers,
        Availability = BookOnixAvailability.Available, Prices = [Price()]
    };
    private static BookOnixCommercialMetadata Commercial() => new() {
        PublishingStatus = BookOnixPublishingStatus.Active,
        SalesRights = [new(BookOnixSalesRightsKind.Exclusive, Countries("PL", "GB"))], Supplies = [Supply()]
    };
    private static XDocument Export(BookOnixCommercialMetadata commercial, DateOnly? publicationDate = null) => XDocument.Load(new MemoryStream(
        BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Commercial = commercial, PublicationDate = publicationDate }, BookOnixTests.TestSchema()).Bytes));

    [Fact]
    public void PricesRetainDecimalTaxBasisMarketAndDatesRegardlessOfHostCulture() {
        CultureInfo previous = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
            var xml = Export(Commercial() with { Supplies = [Supply() with { Prices = [Price() with {
                ValidFrom = new DateOnly(2026, 10, 1), ValidUntil = new DateOnly(2026, 11, 1)
            }] }] }, new DateOnly(2026, 10, 1));
            var price = xml.Descendants(Ns + "Price").Single();
            Assert.Equal("02", price.Element(Ns + "PriceType")!.Value);
            Assert.Equal("12.3400", price.Element(Ns + "PriceAmount")!.Value);
            Assert.Equal("PLN", price.Element(Ns + "CurrencyCode")!.Value);
            Assert.Equal("PL", price.Descendants(Ns + "CountriesIncluded").Single().Value);
            Assert.Equal(new[] { "14", "15" }, price.Descendants(Ns + "PriceDateRole").Select(e => e.Value));
            Assert.Equal(new[] { "20261001", "20261101" }, price.Descendants(Ns + "Date").Select(e => e.Value));
            Assert.Equal("04", xml.Descendants(Ns + "PublishingStatus").Single().Value);
            Assert.Equal("20", xml.Descendants(Ns + "ProductAvailability").Single().Value);
        } finally { CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void WorldwideRightsCanBePartitionedAndAFreeOfferIsExplicit() {
        var world = new BookOnixTerritory { Worldwide = true };
        var xml = Export(Commercial() with {
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, world with { ExcludedCountries = ["US"] }),
                new(BookOnixSalesRightsKind.NonExclusive, Countries("US"))],
            Supplies = [Supply() with { Territory = world, Prices = [], Unpriced = BookOnixUnpricedKind.Free }]
        });
        Assert.Equal("US", xml.Descendants(Ns + "CountriesExcluded").Single().Value);
        Assert.Equal(new[] { "01", "02" }, xml.Descendants(Ns + "SalesRightsType").Select(e => e.Value));
        Assert.Empty(xml.Descendants(Ns + "PriceAmount"));
        Assert.Equal("01", xml.Descendants(Ns + "UnpricedItemType").Single().Value);
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("lowercase")]
    [InlineData("mixed")]
    [InlineData("exclusion")]
    [InlineData("duplicate")]
    [InlineData("overlap")]
    [InlineData("denied")]
    [InlineData("market")]
    [InlineData("price")]
    [InlineData("world-price")]
    public void TerritoryErrorsNeverBroadenRightsOrSupply(string kind) {
        var commercial = Commercial();
        commercial = kind switch {
            "empty" => commercial with { SalesRights = [new(BookOnixSalesRightsKind.Exclusive, Countries())] },
            "lowercase" => commercial with { SalesRights = [new(BookOnixSalesRightsKind.Exclusive, Countries("pl"))] },
            "mixed" => commercial with { SalesRights = [new(BookOnixSalesRightsKind.Exclusive, Countries("PL") with { Worldwide = true })] },
            "exclusion" => commercial with { SalesRights = [new(BookOnixSalesRightsKind.Exclusive, Countries("PL") with { ExcludedCountries = ["US"] })] },
            "duplicate" => commercial with { SalesRights = [new(BookOnixSalesRightsKind.Exclusive, Countries("PL", "PL"))] },
            "overlap" => commercial with { SalesRights = [new(BookOnixSalesRightsKind.Exclusive, new() { Worldwide = true }), new(BookOnixSalesRightsKind.RightsNotHeld, Countries("PL"))] },
            "denied" => commercial with { SalesRights = [new(BookOnixSalesRightsKind.RightsNotHeld, Countries("PL"))] },
            "market" => commercial with { Supplies = [Supply() with { Territory = new() { Worldwide = true } }] },
            "price" => commercial with { Supplies = [Supply() with { Prices = [Price() with { Territory = Countries("GB") }] }] },
            _ => commercial with { Supplies = [Supply() with { Prices = [Price() with { Territory = new() { Worldwide = true } }] }] }
        };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Commercial = commercial }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("both")]
    [InlineData("zero")]
    [InlineData("negative")]
    [InlineData("currency")]
    [InlineData("dates")]
    public void InvalidPricingDoesNotAcquireAnImplicitMeaning(string kind) {
        BookOnixSupply supply = kind switch {
            "missing" => Supply() with { Prices = [] },
            "both" => Supply() with { Unpriced = BookOnixUnpricedKind.Free },
            "zero" => Supply() with { Prices = [Price() with { Amount = 0 }] },
            "negative" => Supply() with { Prices = [Price() with { Amount = -1 }] },
            "currency" => Supply() with { Prices = [Price() with { CurrencyCode = "pln" }] },
            _ => Supply() with { Prices = [Price() with { ValidFrom = new DateOnly(2026, 11, 1), ValidUntil = new DateOnly(2026, 10, 1) }] }
        };
        Assert.ThrowsAny<ArgumentException>(() => Export(Commercial() with { Supplies = [supply] }));
    }

    [Fact]
    public void AvailabilityDatesAndPublicationDatesHaveSeparateContracts() {
        var forthcoming = Commercial() with { PublishingStatus = BookOnixPublishingStatus.Forthcoming,
            Supplies = [Supply() with { Availability = BookOnixAvailability.NotYetAvailable, ExpectedSupplyDate = new DateOnly(2026, 11, 3) }] };
        Assert.Throws<ArgumentException>(() => Export(forthcoming));
        var xml = Export(forthcoming, new DateOnly(2026, 11, 1));
        Assert.Equal("20261103", xml.Descendants(Ns + "SupplyDate").Single().Element(Ns + "Date")!.Value);
        Assert.Throws<ArgumentException>(() => Export(Commercial() with { PublishingStatus = BookOnixPublishingStatus.Cancelled }, new DateOnly(2026, 11, 1)));
        Assert.Throws<ArgumentException>(() => Export(Commercial() with { Supplies = [Supply() with { Availability = BookOnixAvailability.TemporarilyUnavailable }] }));
        xml = Export(Commercial() with { Supplies = [Supply() with { Availability = BookOnixAvailability.TemporarilyUnavailable, ExpectedSupplyDateUnknown = true }] });
        Assert.Empty(xml.Descendants(Ns + "SupplyDate"));
        Assert.Throws<ArgumentException>(() => Export(Commercial() with { Supplies = [Supply() with { ExpectedSupplyDateUnknown = true }] }));
    }

    [Fact]
    public void WithdrawalCanBeReportedAfterRightsAreLost() {
        var xml = Export(Commercial() with {
            SalesRights = [new(BookOnixSalesRightsKind.RightsNotHeld, Countries("PL"))],
            Supplies = [Supply() with { Availability = BookOnixAvailability.Withdrawn, Prices = [], Unpriced = BookOnixUnpricedKind.ContactSupplier }]
        });
        Assert.Equal("06", xml.Descendants(Ns + "SalesRightsType").Single().Value);
        Assert.Equal("46", xml.Descendants(Ns + "ProductAvailability").Single().Value);
    }

    [Fact]
    public void CommercialRecordCountsAreBounded() {
        Assert.Throws<ArgumentException>(() => Export(Commercial() with { Supplies = Enumerable.Repeat(Supply(), 33).ToArray() }));
        Assert.Throws<ArgumentException>(() => Export(Commercial() with { SalesRights = Enumerable.Repeat(new BookOnixSalesRights(BookOnixSalesRightsKind.Exclusive, Countries("PL")), 33).ToArray() }));
        Assert.Throws<ArgumentException>(() => Export(Commercial() with { Supplies = [Supply() with { Prices = Enumerable.Repeat(Price(), 17).ToArray() }] }));
    }
}
