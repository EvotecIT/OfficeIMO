using OfficeIMO.Workflows;

internal static class TaxFixtures {
    // Synthetic serialization examples, not jurisdictional tax-rate recommendations.
    internal static BookOnixPrice[] Prices() {
        var prices = new List<BookOnixPrice>();
        foreach (BookOnixTaxRateCode code in Enum.GetValues<BookOnixTaxRateCode>()) {
            decimal rate = code switch {
                BookOnixTaxRateCode.Higher => 20m, BookOnixTaxRateCode.Standard => 10m,
                BookOnixTaxRateCode.SuperLow => 2m, BookOnixTaxRateCode.Zero => 0m, _ => 5m
            };
            prices.Add(new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, Amount = 10m + rate / 10m,
                CurrencyCode = "PLN", Taxes = [new() { Type = BookOnixTaxType.ValueAdded,
                    RateCode = code, RatePercent = rate, TaxableAmount = 10.000m, Amount = rate / 10m }] });
        }
        prices.Add(new() { Kind = BookOnixPriceKind.FixedIncludingTax, Amount = 11.00m, CurrencyCode = "PLN",
            Taxes = [new() { Type = BookOnixTaxType.Sales, RatePercent = 5.00m, TaxableAmount = 10.000m, Amount = 0.5000m,
                PricePartDescription = "Digital text & illustrations" },
                new() { Type = BookOnixTaxType.Environmental, Amount = 0.5000m, PricePartDescription = "Separate component" }] });
        prices.Add(new() { Kind = BookOnixPriceKind.PublisherRetailIncludingTax, Amount = 12m, CurrencyCode = "PLN",
            Taxes = [new() { Type = BookOnixTaxType.ValueAdded, RatePercent = 20.00m }] });
        prices.Add(new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, Amount = 10m, CurrencyCode = "PLN", TaxExempt = true });
        return prices.ToArray();
    }
}
