using OfficeIMO.Workflows;

internal static class CommercialFixtures {
    internal static BookOnixCommercialMetadata Create(string profile) {
        var world = new BookOnixTerritory { Worldwide = true };
        var poland = new BookOnixTerritory { Countries = ["PL"] };
        var britain = new BookOnixTerritory { Countries = ["GB"] };
        BookOnixSupply supply = new() {
            Territory = world, SupplierName = "Example Digital Supplier", SupplierRole = BookOnixSupplierRole.PublisherToCustomers,
            Availability = BookOnixAvailability.Available, Unpriced = BookOnixUnpricedKind.Free
        };
        var commercial = new BookOnixCommercialMetadata {
            PublishingStatus = BookOnixPublishingStatus.Active,
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, world)], Supplies = [supply]
        };
        return profile switch {
            "early" => commercial with {
                PublishingStatus = BookOnixPublishingStatus.Forthcoming,
                SalesRights = [new(BookOnixSalesRightsKind.Exclusive, world with { ExcludedCountries = ["US"] }),
                    new(BookOnixSalesRightsKind.RightsNotHeld, new() { Countries = ["US"] })],
                Supplies = [supply with { Territory = world with { ExcludedCountries = ["US"] },
                    SupplierRole = BookOnixSupplierRole.ExclusiveDistributorToResellers,
                    Availability = BookOnixAvailability.NotYetAvailable, ExpectedSupplyDate = new DateOnly(2026, 10, 6),
                    Unpriced = BookOnixUnpricedKind.ToBeAnnounced }]
            },
            "advance" => commercial with {
                PublishingStatus = BookOnixPublishingStatus.Forthcoming,
                SalesRights = [new(BookOnixSalesRightsKind.Exclusive, poland), new(BookOnixSalesRightsKind.NonExclusive, britain)],
                Supplies = [supply with { Territory = poland, Unpriced = null,
                    Availability = BookOnixAvailability.NotYetAvailable, ExpectedSupplyDate = new DateOnly(2026, 10, 6),
                    Prices = [new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, Amount = 34.9900m, CurrencyCode = "PLN",
                        ValidFrom = new DateOnly(2026, 10, 5), ValidUntil = new DateOnly(2026, 11, 5) }] },
                    supply with { Territory = britain, Unpriced = null,
                        SupplierRole = BookOnixSupplierRole.NonExclusiveDistributorToResellers,
                        Availability = BookOnixAvailability.TemporarilyUnavailable, ExpectedSupplyDateUnknown = true,
                        Prices = [new() { Kind = BookOnixPriceKind.RecommendedExcludingTax, Amount = 7.99m, CurrencyCode = "GBP" }] }]
            },
            "priced" => commercial with { Supplies = [supply with { Unpriced = null,
                Prices = Enum.GetValues<BookOnixPriceKind>().Select(kind => new BookOnixPrice {
                    Kind = kind, Amount = 12.50m, CurrencyCode = "PLN", Territory = poland
                }).ToArray() }] },
            "discounted" => commercial with { Supplies = [supply with { Territory = britain, SupplierRole = BookOnixSupplierRole.PublisherToResellers, Unpriced = null, Prices = DiscountFixtures.Prices() }] },
            "taxed" => commercial with { Supplies = [supply with { Territory = poland, Unpriced = null, Prices = TaxFixtures.Prices() }] },
            "withdrawn" => commercial with {
                PublishingStatus = BookOnixPublishingStatus.Withdrawn,
                SalesRights = [new(BookOnixSalesRightsKind.RightsNotHeld, world)],
                Supplies = [supply with { Availability = BookOnixAvailability.Withdrawn, Unpriced = BookOnixUnpricedKind.ContactSupplier }]
            },
            _ => commercial
        };
    }
}
