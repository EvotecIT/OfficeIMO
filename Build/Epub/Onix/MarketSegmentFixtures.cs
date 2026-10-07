using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class MarketSegmentFixtures {
    internal static BookOnixCommercialMetadata Create() {
        var us = new BookOnixTerritory { Countries = ["US"] };
        return new() {
            PublishingStatus = BookOnixPublishingStatus.Active,
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, us)],
            Supplies = [new() {
                Territory = us, MarketReference = "segmented-us", SupplierName = "Example Supplier",
                SupplierRole = BookOnixSupplierRole.PublisherToResellers, Availability = BookOnixAvailability.Available,
                Restrictions = [new(BookOnixSalesRestrictionKind.ExceptOnlineRetail)],
                MarketSegments = [
                    new() { Territory = new() { Regions = ["US-CA"] }, Restrictions = [new(BookOnixSalesRestrictionKind.LibrariesOnly)] },
                    new() { Territory = us with { ExcludedRegions = ["US-CA"] }, Restrictions = [new(BookOnixSalesRestrictionKind.ExceptLibraries)] }
                ],
                Prices = [new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, Amount = 12.50m, CurrencyCode = "USD" }]
            }]
        };
    }

    internal static void Verify(BookOnixExportResult result, XmlSchemaSet schemas, string output) {
        XNamespace ns = BookProject.OnixNamespace;
        if (!BookOnixMessage.Create([result], schemas).Bytes.SequenceEqual(result.Bytes))
            throw new InvalidDataException("Market segmentation changed on record composition.");
        var original = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(ns + "ProductSupply").Single();
        var update = BookOnixMessage.CreateBlockUpdates([new(result) { ReplaceMarketReferences = ["segmented-us"] }], schemas);
        var updated = XDocument.Load(new MemoryStream(update.Bytes)).Descendants(ns + "ProductSupply").Single();
        if (!XNode.DeepEquals(original, updated)) throw new InvalidDataException("Market update changed supply segments.");
        File.WriteAllBytes(Path.Combine(output, "update-market-segments.onix"), update.Bytes);
    }
}
