using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixDiscountCodes(BookOnixPrice price) {
        ArgumentNullException.ThrowIfNull(price.DiscountCodes);
        if (price.DiscountCodes.Count > 16)
            throw new ArgumentException("Supply at most 16 coded discount declarations per price.", nameof(price));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        foreach (BookOnixDiscountCode discount in price.DiscountCodes) {
            ArgumentNullException.ThrowIfNull(discount);
            string scheme = discount.Scheme switch {
                BookOnixDiscountScheme.BicDiscount => "01", BookOnixDiscountScheme.ProprietaryDiscount => "02",
                BookOnixDiscountScheme.Boeksoort => "03", BookOnixDiscountScheme.GermanTerms => "04",
                BookOnixDiscountScheme.ProprietaryCommission => "05", BookOnixDiscountScheme.BicCommission => "06",
                BookOnixDiscountScheme.IsniDiscount => "07", _ => throw new ArgumentOutOfRangeException(nameof(discount.Scheme))
            };
            RequireOnixText(discount.Code, nameof(discount.Code));
            bool proprietary = discount.Scheme is BookOnixDiscountScheme.ProprietaryDiscount or BookOnixDiscountScheme.ProprietaryCommission;
            if (proprietary) RequireOnixText(discount.SchemeName!, nameof(discount.SchemeName));
            else if (discount.SchemeName != null)
                throw new ArgumentException("Only proprietary discount or commission codes carry a scheme name.", nameof(price));
            string? pattern = discount.Scheme switch {
                BookOnixDiscountScheme.BicDiscount or BookOnixDiscountScheme.BicCommission => @"\A[A-Za-z]{5}[A-Za-z0-9]{1,3}\z",
                BookOnixDiscountScheme.IsniDiscount => @"\A[0-9]{15}[0-9X]-[A-Za-z0-9]{1,3}\z",
                _ => null
            };
            if (pattern != null && !Regex.IsMatch(discount.Code, pattern, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1)))
                throw new ArgumentException("The discount code does not match its BIC or ISNI-based structural format.", nameof(price));
            var element = new XElement(ns + "DiscountCoded", new XElement(ns + "DiscountCodeType", scheme));
            if (discount.SchemeName != null) element.Add(new XElement(ns + "DiscountCodeTypeName", discount.SchemeName));
            element.Add(new XElement(ns + "DiscountCode", discount.Code));
            result.Add(element);
        }
        return result;
    }
}
