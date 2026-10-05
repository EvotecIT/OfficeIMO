using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private sealed record OnixCommercialParts(XElement? Status, XElement[] Rights, XElement[] Supplies);

    private static OnixCommercialParts BuildOnixCommercial(BookOnixCommercialMetadata? commercial,
        DateOnly? publicationDate, CancellationToken token) {
        if (commercial == null) return new(null, [], []);
        ArgumentNullException.ThrowIfNull(commercial.SalesRights);
        ArgumentNullException.ThrowIfNull(commercial.Supplies);
        if (commercial.SalesRights.Count > 32 || commercial.Supplies.Count > 32)
            throw new ArgumentException("ONIX export supports at most 32 rights and 32 supply declarations.", nameof(commercial));
        XNamespace ns = OnixNamespace;
        XElement? status = null;
        if (commercial.PublishingStatus is { } publishingStatus) {
            string code = publishingStatus switch {
                BookOnixPublishingStatus.Cancelled => "01", BookOnixPublishingStatus.Forthcoming => "02",
                BookOnixPublishingStatus.PostponedIndefinitely => "03", BookOnixPublishingStatus.Active => "04",
                BookOnixPublishingStatus.OutOfPrint => "07", BookOnixPublishingStatus.Withdrawn => "11",
                _ => throw new ArgumentOutOfRangeException(nameof(commercial.PublishingStatus))
            };
            if (publishingStatus == BookOnixPublishingStatus.Forthcoming && publicationDate == null ||
                publishingStatus is BookOnixPublishingStatus.Cancelled or BookOnixPublishingStatus.PostponedIndefinitely && publicationDate != null)
                throw new ArgumentException("Forthcoming status needs a publication date; cancelled or indefinitely postponed status forbids one.", nameof(commercial));
            status = new XElement(ns + "PublishingStatus", code);
        }
        var rights = new List<XElement>(); var territories = new List<OnixTerritory>(); var grants = new List<OnixTerritory>();
        foreach (BookOnixSalesRights right in commercial.SalesRights) {
            token.ThrowIfCancellationRequested(); ArgumentNullException.ThrowIfNull(right);
            string code = right.Kind switch {
                BookOnixSalesRightsKind.Exclusive => "01", BookOnixSalesRightsKind.NonExclusive => "02",
                BookOnixSalesRightsKind.NotForSale => "03", BookOnixSalesRightsKind.NotForSaleExclusive => "04",
                BookOnixSalesRightsKind.NotForSaleNonExclusive => "05", BookOnixSalesRightsKind.RightsNotHeld => "06",
                _ => throw new ArgumentOutOfRangeException(nameof(right.Kind))
            };
            OnixTerritory territory = ReadOnixTerritory(right.Territory);
            if (territories.Any(existing => existing.Overlaps(territory)))
                throw new ArgumentException("Sales-rights territories must not overlap in this export profile.", nameof(commercial));
            territories.Add(territory);
            if (right.Kind is BookOnixSalesRightsKind.Exclusive or BookOnixSalesRightsKind.NonExclusive) grants.Add(territory);
            rights.Add(new XElement(ns + "SalesRights", new XElement(ns + "SalesRightsType", code), territory.ToXml()));
        }
        var supplies = new List<XElement>();
        foreach (BookOnixSupply supply in commercial.Supplies) {
            token.ThrowIfCancellationRequested(); ArgumentNullException.ThrowIfNull(supply);
            supplies.Add(BuildOnixSupply(supply, grants, token));
        }
        return new(status, rights.ToArray(), supplies.ToArray());
    }

    private static XElement BuildOnixSupply(BookOnixSupply supply, IReadOnlyList<OnixTerritory> grants, CancellationToken token) {
        XNamespace ns = OnixNamespace;
        OnixTerritory market = ReadOnixTerritory(supply.Territory);
        RequireOnixText(supply.SupplierName, nameof(supply.SupplierName));
        string role = supply.SupplierRole switch {
            BookOnixSupplierRole.PublisherToResellers => "01", BookOnixSupplierRole.ExclusiveDistributorToResellers => "02",
            BookOnixSupplierRole.NonExclusiveDistributorToResellers => "03", BookOnixSupplierRole.Wholesaler => "04",
            BookOnixSupplierRole.Retailer => "08", BookOnixSupplierRole.PublisherToCustomers => "09",
            BookOnixSupplierRole.ExclusiveDistributorToCustomers => "10", BookOnixSupplierRole.NonExclusiveDistributorToCustomers => "11",
            _ => throw new ArgumentOutOfRangeException(nameof(supply.SupplierRole))
        };
        string availability = supply.Availability switch {
            BookOnixAvailability.NotYetAvailable => "10", BookOnixAvailability.Available => "20",
            BookOnixAvailability.TemporarilyUnavailable => "30", BookOnixAvailability.Unavailable => "40",
            BookOnixAvailability.Withdrawn => "46", _ => throw new ArgumentOutOfRangeException(nameof(supply.Availability))
        };
        bool expected = supply.Availability is BookOnixAvailability.NotYetAvailable or BookOnixAvailability.TemporarilyUnavailable;
        if (expected ? supply.ExpectedSupplyDate.HasValue == supply.ExpectedSupplyDateUnknown : supply.ExpectedSupplyDate.HasValue || supply.ExpectedSupplyDateUnknown)
            throw new ArgumentException("Expected supply needs a date or explicit unknown-date assertion; other availability states cannot carry an expected date.", nameof(supply));
        if (supply.Availability is not (BookOnixAvailability.Unavailable or BookOnixAvailability.Withdrawn) && !OnixGrantsCover(market, grants))
            throw new ArgumentException("Available or expected supply must lie within explicitly declared for-sale rights.", nameof(supply));
        ArgumentNullException.ThrowIfNull(supply.Prices);
        if (supply.Prices.Count > 16 || supply.Unpriced.HasValue == (supply.Prices.Count != 0))
            throw new ArgumentException("Supply needs 1-16 prices or one explicit unpriced reason, not both.", nameof(supply));
        var detail = new XElement(ns + "SupplyDetail", new XElement(ns + "Supplier",
            new XElement(ns + "SupplierRole", role), new XElement(ns + "SupplierName", supply.SupplierName)),
            new XElement(ns + "ProductAvailability", availability));
        if (supply.ExpectedSupplyDate is { } expectedDate) detail.Add(OnixCommercialDate("SupplyDate", "SupplyDateRole", "08", expectedDate));
        if (supply.Unpriced is { } unpriced) {
            string code = unpriced switch {
                BookOnixUnpricedKind.Free => "01", BookOnixUnpricedKind.ToBeAnnounced => "02", BookOnixUnpricedKind.ContactSupplier => "04",
                _ => throw new ArgumentOutOfRangeException(nameof(supply.Unpriced))
            };
            detail.Add(new XElement(ns + "UnpricedItemType", code));
        } else foreach (BookOnixPrice price in supply.Prices) {
            token.ThrowIfCancellationRequested(); ArgumentNullException.ThrowIfNull(price);
            detail.Add(BuildOnixPrice(price, market));
        }
        return new XElement(ns + "ProductSupply", new XElement(ns + "Market", market.ToXml()), detail);
    }

    private static XElement BuildOnixPrice(BookOnixPrice price, OnixTerritory market) {
        XNamespace ns = OnixNamespace;
        string type = price.Kind switch {
            BookOnixPriceKind.RecommendedExcludingTax => "01", BookOnixPriceKind.RecommendedIncludingTax => "02",
            BookOnixPriceKind.FixedExcludingTax => "03", BookOnixPriceKind.FixedIncludingTax => "04",
            BookOnixPriceKind.SupplierNetExcludingTax => "05", BookOnixPriceKind.PublisherRetailExcludingTax => "41",
            BookOnixPriceKind.PublisherRetailIncludingTax => "42", _ => throw new ArgumentOutOfRangeException(nameof(price.Kind))
        };
        if (price.Amount <= 0) throw new ArgumentOutOfRangeException(nameof(price.Amount), "Prices must be positive; use Unpriced = Free for free supply.");
        if (price.CurrencyCode == null || price.CurrencyCode.Length != 3 || price.CurrencyCode.Any(c => c < 'A' || c > 'Z'))
            throw new ArgumentException("Supply an uppercase three-letter ONIX list 96 currency code.", nameof(price.CurrencyCode));
        if (price.ValidFrom > price.ValidUntil) throw new ArgumentException("Price dates cannot run backwards.", nameof(price));
        OnixTerritory territory = price.Territory == null ? market : ReadOnixTerritory(price.Territory);
        if (!market.Covers(territory)) throw new ArgumentException("A price territory cannot extend outside its supply market.", nameof(price));
        var result = new XElement(ns + "Price", new XElement(ns + "PriceType", type),
            new XElement(ns + "PriceAmount", price.Amount.ToString(CultureInfo.InvariantCulture)),
            new XElement(ns + "CurrencyCode", price.CurrencyCode), territory.ToXml());
        if (price.ValidFrom is { } from) result.Add(OnixCommercialDate("PriceDate", "PriceDateRole", "14", from));
        if (price.ValidUntil is { } until) result.Add(OnixCommercialDate("PriceDate", "PriceDateRole", "15", until));
        return result;
    }

    private static XElement OnixCommercialDate(string composite, string roleElement, string role, DateOnly date) {
        XNamespace ns = OnixNamespace;
        return new XElement(ns + composite, new XElement(ns + roleElement, role),
            new XElement(ns + "Date", new XAttribute("dateformat", "00"), date.ToString("yyyyMMdd", CultureInfo.InvariantCulture)));
    }
}
