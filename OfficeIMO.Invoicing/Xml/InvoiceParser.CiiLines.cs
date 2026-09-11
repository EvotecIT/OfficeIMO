using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceParser {
    private static InvoiceLine CiiLine(InvoiceXmlReadContext c, XElement element, string currency) {
        XElement? document = c.Child(element, Ram + "AssociatedDocumentLineDocument"), product = c.Child(element, Ram + "SpecifiedTradeProduct");
        XElement? agreement = c.Child(element, Ram + "SpecifiedLineTradeAgreement"), delivery = c.Child(element, Ram + "SpecifiedLineTradeDelivery"), settlement = c.Child(element, Ram + "SpecifiedLineTradeSettlement");
        XElement? netPrice = c.Child(agreement, Ram + "NetPriceProductTradePrice"), quantity = c.Child(delivery, Ram + "BilledQuantity");
        XElement? basis = c.Child(netPrice, Ram + "BasisQuantity");
        var line = new InvoiceLine { Id = c.Required(document, Ram + "LineID"), Note = c.Text(c.Child(document, Ram + "IncludedNote"), Ram + "Content"),
            Name = c.Required(product, Ram + "Name"), Description = c.Text(product, Ram + "Description"),
            StandardItemIdentifier = c.Identifier(c.Child(product, Ram + "GlobalID")), SellerItemIdentifier = c.Text(product, Ram + "SellerAssignedID"), BuyerItemIdentifier = c.Text(product, Ram + "BuyerAssignedID"),
            Quantity = c.RequiredDecimal(quantity), UnitCode = c.Attribute(quantity, "unitCode") ?? string.Empty,
            UnitPrice = c.RequiredMoney(netPrice, Ram + "ChargeAmount", currency), PriceBaseQuantity = c.Decimal(basis) ?? 1m,
            Tax = CiiTaxCategory(c, c.Child(settlement, Ram + "ApplicableTradeTax")), Period = CiiPeriod(c, settlement),
            OrderLineReference = c.Text(c.Child(agreement, Ram + "BuyerOrderReferencedDocument"), Ram + "LineID"),
            AccountingReference = c.Text(c.Child(settlement, Ram + "ReceivableSpecifiedTradeAccountingAccount"), Ram + "ID"),
            DeclaredNetAmount = c.RequiredMoney(c.Child(settlement, Ram + "SpecifiedTradeSettlementLineMonetarySummation"), Ram + "LineTotalAmount", currency),
            OriginCountryCode = c.Text(c.Child(product, Ram + "OriginTradeCountry"), Ram + "ID") };
        string? priceUnit = c.Attribute(basis, "unitCode");
        if (priceUnit != null && priceUnit != line.UnitCode) c.Loss(basis!, "Price base unit differs from billed quantity unit.");
        XElement? gross = c.Child(agreement, Ram + "GrossPriceProductTradePrice");
        line.GrossPrice = c.Money(gross, Ram + "ChargeAmount", currency);
        XElement? grossBasis = c.Child(gross, Ram + "BasisQuantity");
        decimal? grossQuantity = c.Decimal(grossBasis);
        string? grossUnit = c.Attribute(grossBasis, "unitCode");
        if (grossQuantity.HasValue && grossQuantity != line.PriceBaseQuantity || grossUnit != null && grossUnit != line.UnitCode)
            c.Loss(grossBasis!, "Gross and net price basis differ.");
        XElement? discount = c.Child(gross, Ram + "AppliedTradeAllowanceCharge");
        if (discount != null) {
            if (c.Boolean(c.Child(c.Child(discount, Ram + "ChargeIndicator"), Udt + "Indicator"))) c.Loss(discount, "Gross-price charges are outside the supported price-discount mapping.");
            line.PriceDiscount = c.Money(discount, Ram + "ActualAmount", currency);
        }
        foreach (XElement adjustment in c.Children(settlement, Ram + "SpecifiedTradeAllowanceCharge")) c.AddTo(line.AllowancesAndCharges, CiiAdjustment(c, adjustment, currency, false));
        foreach (XElement classification in c.Children(product, Ram + "DesignatedProductClassification")) {
            XElement? code = c.Child(classification, Ram + "ClassCode");
            c.AddTo(line.Classifications, new InvoiceItemClassification { Value = c.Value(code) ?? string.Empty, ListId = c.Attribute(code, "listID") ?? string.Empty, ListVersion = c.Attribute(code, "listVersionID") });
        }
        foreach (XElement attribute in c.Children(product, Ram + "ApplicableProductCharacteristic"))
            c.AddTo(line.Attributes, new InvoiceItemAttribute { Name = c.Required(attribute, Ram + "Description"), Value = c.Required(attribute, Ram + "Value") });
        XElement? objectReference = c.Child(settlement, Ram + "AdditionalReferencedDocument");
        if (objectReference != null) {
            c.Expected(c.Child(objectReference, Ram + "TypeCode"), "130");
            line.ObjectIdentifier = new InvoiceIdentifier(c.Required(objectReference, Ram + "IssuerAssignedID"), c.Text(objectReference, Ram + "ReferenceTypeCode"));
        }
        return line;
    }
}
