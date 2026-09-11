using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static XElement CiiLine(InvoiceLine line, InvoiceCalculatedLine calculation) => new XElement(Ram + "IncludedSupplyChainTradeLineItem",
        new XElement(Ram + "AssociatedDocumentLineDocument", Text(Ram + "LineID", line.Id), line.Note == null ? null : new XElement(Ram + "IncludedNote", Text(Ram + "Content", line.Note))),
        new XElement(Ram + "SpecifiedTradeProduct", Identifier(Ram + "GlobalID", line.StandardItemIdentifier), Text(Ram + "SellerAssignedID", line.SellerItemIdentifier),
            Text(Ram + "BuyerAssignedID", line.BuyerItemIdentifier), Text(Ram + "Name", line.Name), Text(Ram + "Description", line.Description),
            line.Attributes.Select(attribute => new XElement(Ram + "ApplicableProductCharacteristic", Text(Ram + "Description", attribute.Name), Text(Ram + "Value", attribute.Value))),
            line.Classifications.Select(classification => new XElement(Ram + "DesignatedProductClassification", new XElement(Ram + "ClassCode", new XAttribute("listID", classification.ListId),
                classification.ListVersion == null ? null : new XAttribute("listVersionID", classification.ListVersion), classification.Value))),
            line.OriginCountryCode == null ? null : new XElement(Ram + "OriginTradeCountry", Text(Ram + "ID", line.OriginCountryCode))),
        new XElement(Ram + "SpecifiedLineTradeAgreement",
            line.OrderLineReference == null ? null : new XElement(Ram + "BuyerOrderReferencedDocument", Text(Ram + "LineID", line.OrderLineReference)),
            line.GrossPrice.HasValue ? new XElement(Ram + "GrossPriceProductTradePrice", Text(Ram + "ChargeAmount", Number(line.GrossPrice.Value)),
                new XElement(Ram + "BasisQuantity", new XAttribute("unitCode", line.UnitCode), Number(line.PriceBaseQuantity)),
                line.PriceDiscount.HasValue ? new XElement(Ram + "AppliedTradeAllowanceCharge", new XElement(Ram + "ChargeIndicator", new XElement(Udt + "Indicator", "false")), Text(Ram + "ActualAmount", Number(line.PriceDiscount.Value))) : null) : null,
            new XElement(Ram + "NetPriceProductTradePrice", Text(Ram + "ChargeAmount", Number(line.UnitPrice)),
                new XElement(Ram + "BasisQuantity", new XAttribute("unitCode", line.UnitCode), Number(line.PriceBaseQuantity)))),
        new XElement(Ram + "SpecifiedLineTradeDelivery", new XElement(Ram + "BilledQuantity", new XAttribute("unitCode", line.UnitCode), Number(line.Quantity))),
        new XElement(Ram + "SpecifiedLineTradeSettlement", new XElement(Ram + "ApplicableTradeTax", new XElement(Ram + "TypeCode", "VAT"), Text(Ram + "CategoryCode", line.Tax.Code),
                line.Tax.Rate.HasValue ? Text(Ram + "RateApplicablePercent", Number(line.Tax.Rate.Value)) : null),
            CiiPeriod(line.Period), line.AllowancesAndCharges.Select(item => CiiAdjustment(item, false)),
            new XElement(Ram + "SpecifiedTradeSettlementLineMonetarySummation", CiiAmount("LineTotalAmount", calculation.NetAmount)),
            CiiObjectReference(line.ObjectIdentifier),
            line.AccountingReference == null ? null : new XElement(Ram + "ReceivableSpecifiedTradeAccountingAccount", Text(Ram + "ID", line.AccountingReference))));
}
