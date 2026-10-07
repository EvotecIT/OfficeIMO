using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixTaxes(BookOnixPrice price) {
        ArgumentNullException.ThrowIfNull(price.Taxes);
        if (price.Taxes.Count > 16 || price.TaxExempt && price.Taxes.Count != 0)
            throw new ArgumentException("Supply at most 16 tax components, or TaxExempt, not both.", nameof(price));
        XNamespace ns = OnixNamespace;
        if (price.TaxExempt) return [new XElement(ns + "TaxExempt")];
        if (price.Taxes.Count != 0 && price.Kind is not (BookOnixPriceKind.RecommendedIncludingTax or
            BookOnixPriceKind.FixedIncludingTax or BookOnixPriceKind.PublisherRetailIncludingTax))
            throw new ArgumentException("Tax components describe tax included in a tax-inclusive price.", nameof(price));
        var result = new List<XElement>();
        decimal remaining = price.Amount;
        foreach (BookOnixTax tax in price.Taxes) {
            ArgumentNullException.ThrowIfNull(tax);
            string type = tax.Type switch {
                BookOnixTaxType.ValueAdded => "01", BookOnixTaxType.Sales => "02", BookOnixTaxType.Environmental => "03",
                _ => throw new ArgumentOutOfRangeException(nameof(tax.Type))
            };
            string? rate = tax.RateCode switch {
                null => null, BookOnixTaxRateCode.Higher => "H", BookOnixTaxRateCode.PaidAtSource => "P",
                BookOnixTaxRateCode.Lower => "R", BookOnixTaxRateCode.Standard => "S",
                BookOnixTaxRateCode.SuperLow => "T", BookOnixTaxRateCode.Zero => "Z",
                _ => throw new ArgumentOutOfRangeException(nameof(tax.RateCode))
            };
            if (tax.RatePercent == null && tax.Amount == null || tax.RatePercent is < 0 or > 100 ||
                tax.TaxableAmount is <= 0 || tax.TaxableAmount > price.Amount || tax.Amount < 0 ||
                tax.Amount > price.Amount - (tax.TaxableAmount ?? 0) || tax.Amount > remaining || tax.RateCode == BookOnixTaxRateCode.Zero && (tax.RatePercent > 0 || tax.Amount > 0) ||
                tax.RatePercent == 0 && tax.Amount > 0)
                throw new ArgumentException("Tax components need a rate or amount, valid nonnegative values, and consistent zero-rate assertions.", nameof(price));
            remaining -= tax.Amount ?? 0;
            var element = new XElement(ns + "Tax");
            if (tax.PricePartDescription != null) {
                RequireOnixText(tax.PricePartDescription, nameof(tax.PricePartDescription));
                element.Add(new XElement(ns + "PricePartDescription", tax.PricePartDescription));
            }
            element.Add(new XElement(ns + "TaxType", type));
            if (rate != null) element.Add(new XElement(ns + "TaxRateCode", rate));
            if (tax.RatePercent is { } percent) element.Add(new XElement(ns + "TaxRatePercent", percent.ToString(CultureInfo.InvariantCulture)));
            if (tax.TaxableAmount is { } taxable) element.Add(new XElement(ns + "TaxableAmount", taxable.ToString(CultureInfo.InvariantCulture)));
            if (tax.Amount is { } amount) element.Add(new XElement(ns + "TaxAmount", amount.ToString(CultureInfo.InvariantCulture)));
            result.Add(element);
        }
        return result;
    }
}
