using OfficeIMO.Workflows;

internal static class SubterritoryFixtures {
    internal static BookOnixCommercialMetadata Create() => new() {
        PublishingStatus = BookOnixPublishingStatus.Active,
        SalesRights = [
            new(BookOnixSalesRightsKind.Exclusive, new() { Worldwide = true, ExcludedRegions = ["US-CA", "CA-QC"] }),
            new(BookOnixSalesRightsKind.NonExclusive, new() { Regions = ["CA-QC", "US-CA"] })],
        Supplies = [new() {
            Territory = new() { Countries = ["GB", "US"], Regions = ["CA-QC"], ExcludedRegions = ["US-NY"] },
            SupplierName = "Example Regional Supplier", SupplierRole = BookOnixSupplierRole.PublisherToCustomers,
            Availability = BookOnixAvailability.Available,
            Prices = [new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, Amount = 12.50m, CurrencyCode = "USD",
                Territory = new() { Regions = ["US-CA"] } }]
        }]
    };
}
