using OfficeIMO.Workflows;

internal static class DiscountFixtures {
    internal static BookOnixPrice[] Prices() {
        var prices = Enum.GetValues<BookOnixDiscountKind>().Select(kind => new BookOnixPrice {
            Kind = BookOnixPriceKind.RecommendedExcludingTax, Amount = 10.00m, CurrencyCode = "GBP",
            Discounts = [new() { Kind = kind, MinimumQuantity = 1, MaximumQuantity = 9, Percent = 10.00m, Amount = 1.000m },
                new() { Kind = kind, MinimumQuantity = 10, Percent = 20.00m }]
        }).ToList();
        prices.Add(new() { Kind = BookOnixPriceKind.RecommendedExcludingTax, Amount = 10.00m, CurrencyCode = "GBP",
            Discounts = [new() { Kind = BookOnixDiscountKind.Rising, Amount = 2.5000m }] });
        prices.Add(new() { Kind = BookOnixPriceKind.RecommendedExcludingTax, Amount = 10.00m, CurrencyCode = "GBP",
            Discounts = [new() { Kind = BookOnixDiscountKind.Rising, Percent = 0, Amount = 0 }] });
        return prices.ToArray();
    }
}
