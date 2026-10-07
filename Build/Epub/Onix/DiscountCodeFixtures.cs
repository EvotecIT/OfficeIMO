using OfficeIMO.Workflows;

internal static class DiscountCodeFixtures {
    // Synthetic scheme declarations; no prefix, identifier allocation, or partner agreement is claimed.
    internal static BookOnixPrice[] Prices() => Enum.GetValues<BookOnixDiscountScheme>().Select(scheme => new BookOnixPrice {
        Kind = scheme is BookOnixDiscountScheme.ProprietaryCommission or BookOnixDiscountScheme.BicCommission
            ? BookOnixPriceKind.PublisherRetailExcludingTax : BookOnixPriceKind.RecommendedExcludingTax,
        Amount = 10.00m, CurrencyCode = "GBP",
        Discounts = scheme == BookOnixDiscountScheme.ProprietaryDiscount
            ? [new() { Kind = BookOnixDiscountKind.Rising, Percent = 10.00m }] : [],
        DiscountCodes = [new() {
            Scheme = scheme,
            SchemeName = scheme is BookOnixDiscountScheme.ProprietaryDiscount or BookOnixDiscountScheme.ProprietaryCommission ? "Example & partners" : null,
            Code = scheme switch {
                BookOnixDiscountScheme.BicDiscount or BookOnixDiscountScheme.BicCommission => "ABCDE12",
                BookOnixDiscountScheme.IsniDiscount => "0000000121032683-A1",
                BookOnixDiscountScheme.Boeksoort => "A", BookOnixDiscountScheme.GermanTerms => "1",
                _ => "trade-A"
            }
        }]
    }).ToArray();
}
