using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceParser {
    private static InvoiceLine UblLine(InvoiceXmlReadContext c, XElement element, string currency, bool credit) {
        XElement? quantity = c.Child(element, Cbc + (credit ? "CreditedQuantity" : "InvoicedQuantity")), price = c.Child(element, Cac + "Price"), item = c.Child(element, Cac + "Item");
        XElement? basis = c.Child(price, Cbc + "BaseQuantity");
        var line = new InvoiceLine { Id = c.Required(element, Cbc + "ID"), Note = c.Text(element, Cbc + "Note"), Quantity = c.RequiredDecimal(quantity),
            UnitCode = c.Attribute(quantity, "unitCode") ?? string.Empty, DeclaredNetAmount = c.RequiredMoney(element, Cbc + "LineExtensionAmount", currency, true),
            AccountingReference = c.Text(element, Cbc + "AccountingCost"), OrderLineReference = c.Text(c.Child(element, Cac + "OrderLineReference"), Cbc + "LineID"),
            Name = c.Required(item, Cbc + "Name"), Description = c.Text(item, Cbc + "Description"),
            SellerItemIdentifier = c.Text(c.Child(item, Cac + "SellersItemIdentification"), Cbc + "ID"), BuyerItemIdentifier = c.Text(c.Child(item, Cac + "BuyersItemIdentification"), Cbc + "ID"),
            StandardItemIdentifier = c.Identifier(c.Child(c.Child(item, Cac + "StandardItemIdentification"), Cbc + "ID")),
            OriginCountryCode = c.Text(c.Child(item, Cac + "OriginCountry"), Cbc + "IdentificationCode"), Tax = UblTaxCategory(c, c.Child(item, Cac + "ClassifiedTaxCategory")),
            UnitPrice = c.RequiredMoney(price, Cbc + "PriceAmount", currency, true), PriceBaseQuantity = c.Decimal(basis) ?? 1m };
        line.Period = UblPeriod(c, element, out string? code);
        if (code != null) c.Loss(element, "Line tax-point codes are outside the supported period mapping.");
        string? priceUnit = c.Attribute(basis, "unitCode");
        if (priceUnit != null && priceUnit != line.UnitCode) c.Loss(basis!, "Price base unit differs from invoiced quantity unit.");
        XElement? objectReference = c.Child(element, Cac + "DocumentReference");
        if (objectReference != null) {
            c.Expected(c.Child(objectReference, Cbc + "DocumentTypeCode"), "130");
            line.ObjectIdentifier = c.Identifier(c.Child(objectReference, Cbc + "ID"));
        }
        foreach (XElement adjustment in c.Children(element, Cac + "AllowanceCharge")) line.AllowancesAndCharges.Add(UblAdjustment(c, adjustment, currency, false));
        XElement? discount = c.Child(price, Cac + "AllowanceCharge");
        if (discount != null) {
            if (c.Boolean(c.Child(discount, Cbc + "ChargeIndicator"))) c.Loss(discount, "Item price charges are outside the supported price-discount mapping.");
            line.PriceDiscount = c.Money(discount, Cbc + "Amount", currency, true); line.GrossPrice = c.Money(discount, Cbc + "BaseAmount", currency, true);
        }
        foreach (XElement classification in c.Children(item, Cac + "CommodityClassification")) {
            XElement? value = c.Child(classification, Cbc + "ItemClassificationCode");
            line.Classifications.Add(new InvoiceItemClassification { Value = c.Value(value) ?? string.Empty, ListId = c.Attribute(value, "listID") ?? string.Empty, ListVersion = c.Attribute(value, "listVersionID") });
        }
        foreach (XElement attribute in c.Children(item, Cac + "AdditionalItemProperty")) line.Attributes.Add(new InvoiceItemAttribute { Name = c.Required(attribute, Cbc + "Name"), Value = c.Required(attribute, Cbc + "Value") });
        return line;
    }
}
