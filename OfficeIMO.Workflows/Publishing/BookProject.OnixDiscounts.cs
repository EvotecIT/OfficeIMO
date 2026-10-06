using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixDiscounts(BookOnixPrice price) {
        ArgumentNullException.ThrowIfNull(price.Discounts);
        if (price.Discounts.Count > 16)
            throw new ArgumentException("Supply at most 16 discount declarations per price.", nameof(price));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        foreach (BookOnixDiscount discount in price.Discounts) {
            ArgumentNullException.ThrowIfNull(discount);
            string kind = discount.Kind switch {
                BookOnixDiscountKind.Rising => "01", BookOnixDiscountKind.RisingCumulative => "02",
                BookOnixDiscountKind.Progressive => "03", BookOnixDiscountKind.ProgressiveCumulative => "04",
                _ => throw new ArgumentOutOfRangeException(nameof(discount.Kind))
            };
            if (discount.Percent == null && discount.Amount == null || discount.Percent is < 0 or > 100 ||
                discount.Amount < 0 || discount.Amount > price.Amount || discount.Percent == 0 && discount.Amount > 0)
                throw new ArgumentException("Discounts need a percentage or amount, valid nonnegative values, and consistent zero assertions.", nameof(price));
            if (discount.MinimumQuantity is <= 0 || discount.MaximumQuantity.HasValue && !discount.MinimumQuantity.HasValue ||
                discount.MaximumQuantity < discount.MinimumQuantity)
                throw new ArgumentException("Discount quantity ranges need a positive minimum and an optional maximum at least as large.", nameof(price));
            var element = new XElement(ns + "Discount", new XElement(ns + "DiscountType", kind));
            if (discount.MinimumQuantity is { } minimum) element.Add(new XElement(ns + "Quantity", minimum));
            if (discount.MaximumQuantity is { } maximum) element.Add(new XElement(ns + "ToQuantity", maximum));
            if (discount.Percent is { } percent) element.Add(new XElement(ns + "DiscountPercent", percent.ToString(CultureInfo.InvariantCulture)));
            if (discount.Amount is { } amount) element.Add(new XElement(ns + "DiscountAmount", amount.ToString(CultureInfo.InvariantCulture)));
            result.Add(element);
        }
        return result;
    }
}
