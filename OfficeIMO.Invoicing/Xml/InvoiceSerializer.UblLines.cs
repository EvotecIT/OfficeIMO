using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static XElement UblLine(InvoiceLine line, InvoiceCalculatedLine calculation, string currency, bool credit) => new XElement(Cac + (credit ? "CreditNoteLine" : "InvoiceLine"),
        Text(Cbc + "ID", line.Id), Text(Cbc + "Note", line.Note),
        new XElement(Cbc + (credit ? "CreditedQuantity" : "InvoicedQuantity"), new XAttribute("unitCode", line.UnitCode), Number(line.Quantity)),
        UblAmount("LineExtensionAmount", calculation.NetAmount, currency), Text(Cbc + "AccountingCost", line.AccountingReference), UblPeriod(line.Period),
        line.OrderLineReference == null ? null : new XElement(Cac + "OrderLineReference", Text(Cbc + "LineID", line.OrderLineReference)),
        line.ObjectIdentifier == null ? null : new XElement(Cac + "DocumentReference", Identifier(Cbc + "ID", line.ObjectIdentifier), new XElement(Cbc + "DocumentTypeCode", "130")),
        line.AllowancesAndCharges.Select(item => UblAdjustment(item, currency, false)),
        new XElement(Cac + "Item", Text(Cbc + "Description", line.Description), Text(Cbc + "Name", line.Name),
            line.BuyerItemIdentifier == null ? null : new XElement(Cac + "BuyersItemIdentification", Text(Cbc + "ID", line.BuyerItemIdentifier)),
            line.SellerItemIdentifier == null ? null : new XElement(Cac + "SellersItemIdentification", Text(Cbc + "ID", line.SellerItemIdentifier)),
            line.StandardItemIdentifier == null ? null : new XElement(Cac + "StandardItemIdentification", Identifier(Cbc + "ID", line.StandardItemIdentifier)),
            line.OriginCountryCode == null ? null : new XElement(Cac + "OriginCountry", Text(Cbc + "IdentificationCode", line.OriginCountryCode)),
            line.Classifications.Select(classification => new XElement(Cac + "CommodityClassification", new XElement(Cbc + "ItemClassificationCode", new XAttribute("listID", classification.ListId),
                classification.ListVersion == null ? null : new XAttribute("listVersionID", classification.ListVersion), classification.Value))),
            UblTaxCategory("ClassifiedTaxCategory", line.Tax.Code, line.Tax.Rate),
            line.Attributes.Select(attribute => new XElement(Cac + "AdditionalItemProperty", Text(Cbc + "Name", attribute.Name), Text(Cbc + "Value", attribute.Value)))),
        new XElement(Cac + "Price", new XElement(Cbc + "PriceAmount", new XAttribute("currencyID", currency), Number(line.UnitPrice)),
            new XElement(Cbc + "BaseQuantity", new XAttribute("unitCode", line.UnitCode), Number(line.PriceBaseQuantity)),
            line.GrossPrice.HasValue ? new XElement(Cac + "AllowanceCharge", new XElement(Cbc + "ChargeIndicator", "false"),
                new XElement(Cbc + "Amount", new XAttribute("currencyID", currency), Number(line.PriceDiscount ?? 0m)),
                new XElement(Cbc + "BaseAmount", new XAttribute("currencyID", currency), Number(line.GrossPrice.Value))) : null));
}
