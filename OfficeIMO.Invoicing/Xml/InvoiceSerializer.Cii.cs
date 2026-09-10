using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static XDocument WriteCii(Invoice invoice, InvoiceCalculation calculation, InvoiceXmlOptions options) => new XDocument(
        new XElement(Rsm + "CrossIndustryInvoice", new XAttribute(XNamespace.Xmlns + "rsm", Rsm), new XAttribute(XNamespace.Xmlns + "ram", Ram),
            new XAttribute(XNamespace.Xmlns + "udt", Udt), new XAttribute(XNamespace.Xmlns + "qdt", Qdt),
            new XElement(Rsm + "ExchangedDocumentContext",
                invoice.BusinessProcessId == null ? null : new XElement(Ram + "BusinessProcessSpecifiedDocumentContextParameter", Text(Ram + "ID", invoice.BusinessProcessId)),
                new XElement(Ram + "GuidelineSpecifiedDocumentContextParameter", new XElement(Ram + "ID", InvoiceProfiles.GetGuidelineId(options.Profile)))),
            new XElement(Rsm + "ExchangedDocument", new XElement(Ram + "ID", invoice.Number), new XElement(Ram + "TypeCode", invoice.TypeCode),
                CiiDate("IssueDateTime", invoice.IssueDate), invoice.Notes.Select(note => new XElement(Ram + "IncludedNote", Text(Ram + "Content", note.Text), Text(Ram + "SubjectCode", note.SubjectCode)))),
            new XElement(Rsm + "SupplyChainTradeTransaction",
                invoice.Lines.Select((line, index) => CiiLine(line, calculation.Lines[index])),
                CiiAgreement(invoice), CiiDelivery(invoice), CiiSettlement(invoice, calculation))));

    private static XElement CiiAgreement(Invoice invoice) => new XElement(Ram + "ApplicableHeaderTradeAgreement",
        Text(Ram + "BuyerReference", invoice.BuyerReference), CiiParty("SellerTradeParty", invoice.Seller), CiiParty("BuyerTradeParty", invoice.Buyer),
        invoice.TaxRepresentative == null ? null : CiiParty("SellerTaxRepresentativeTradeParty", invoice.TaxRepresentative),
        CiiReference("SellerOrderReferencedDocument", invoice.SalesOrderReference), CiiReference("BuyerOrderReferencedDocument", invoice.PurchaseOrderReference),
        CiiReference("ContractReferencedDocument", invoice.ContractReference),
        invoice.SupportingDocuments.Select(document => new XElement(Ram + "AdditionalReferencedDocument",
            Text(Ram + "IssuerAssignedID", document.Reference), Text(Ram + "URIID", document.ExternalUri), new XElement(Ram + "TypeCode", "916"), Text(Ram + "Name", document.Description),
            document.Data == null ? null : new XElement(Ram + "AttachmentBinaryObject", new XAttribute("mimeCode", document.MimeType!), new XAttribute("filename", document.FileName!), Convert.ToBase64String(document.Data)))),
        CiiReference("AdditionalReferencedDocument", invoice.TenderReference, "50"), CiiObjectReference(invoice.ObjectIdentifier),
        invoice.ProjectReference == null ? null : new XElement(Ram + "SpecifiedProcuringProject", Text(Ram + "ID", invoice.ProjectReference), new XElement(Ram + "Name", invoice.ProjectReference)));

    private static XElement CiiDelivery(Invoice invoice) => new XElement(Ram + "ApplicableHeaderTradeDelivery",
        invoice.Delivery == null || invoice.Delivery.Name == null && invoice.Delivery.Address == null && invoice.Delivery.LocationIdentifier == null ? null :
            new XElement(Ram + "ShipToTradeParty", Identifier(Ram + (invoice.Delivery.LocationIdentifier?.SchemeId == null ? "ID" : "GlobalID"), invoice.Delivery.LocationIdentifier),
                Text(Ram + "Name", invoice.Delivery.Name), invoice.Delivery.Address == null ? null : CiiAddress(invoice.Delivery.Address)),
        invoice.Delivery?.Date == null ? null : new XElement(Ram + "ActualDeliverySupplyChainEvent", CiiDate("OccurrenceDateTime", invoice.Delivery.Date)),
        CiiReference("DespatchAdviceReferencedDocument", invoice.DespatchAdviceReference), CiiReference("ReceivingAdviceReferencedDocument", invoice.ReceivingAdviceReference));

    private static XElement CiiSettlement(Invoice invoice, InvoiceCalculation calculation) => new XElement(Ram + "ApplicableHeaderTradeSettlement",
        Text(Ram + "CreditorReferenceID", invoice.Payment?.CreditorIdentifier), Text(Ram + "PaymentReference", invoice.Payment?.Reference),
        Text(Ram + "TaxCurrencyCode", invoice.TaxCurrency), new XElement(Ram + "InvoiceCurrencyCode", invoice.Currency),
        invoice.Payee == null ? null : CiiParty("PayeeTradeParty", invoice.Payee, includeAddress: false),
        CiiPayment(invoice.Payment), calculation.Taxes.Select((tax, index) => CiiTax(tax, invoice, index == 0)), CiiPeriod(invoice.Period),
        invoice.AllowancesAndCharges.Select(item => CiiAdjustment(item, true)),
        invoice.PaymentTerms == null && invoice.DueDate == null && invoice.Payment?.MandateReference == null ? null :
            new XElement(Ram + "SpecifiedTradePaymentTerms", Text(Ram + "Description", invoice.PaymentTerms), CiiDate("DueDateDateTime", invoice.DueDate), Text(Ram + "DirectDebitMandateID", invoice.Payment?.MandateReference)),
        new XElement(Ram + "SpecifiedTradeSettlementHeaderMonetarySummation", CiiAmount("LineTotalAmount", calculation.LineNetTotal),
            CiiAmount("ChargeTotalAmount", calculation.ChargeTotal), CiiAmount("AllowanceTotalAmount", calculation.AllowanceTotal), CiiAmount("TaxBasisTotalAmount", calculation.TaxExclusiveTotal),
            new XElement(Ram + "TaxTotalAmount", new XAttribute("currencyID", invoice.Currency), Amount(calculation.TaxTotal)),
            invoice.TaxAmountInAccountingCurrency.HasValue ? new XElement(Ram + "TaxTotalAmount", new XAttribute("currencyID", invoice.TaxCurrency!), Amount(invoice.TaxAmountInAccountingCurrency.Value)) : null,
            CiiAmount("RoundingAmount", calculation.RoundingAmount), CiiAmount("GrandTotalAmount", calculation.TaxInclusiveTotal),
            CiiAmount("TotalPrepaidAmount", calculation.PrepaidAmount), CiiAmount("DuePayableAmount", calculation.PayableAmount)),
        invoice.PrecedingInvoices.Select(reference => new XElement(Ram + "InvoiceReferencedDocument", Text(Ram + "IssuerAssignedID", reference.Number), CiiDate("FormattedIssueDateTime", reference.IssueDate, true))),
        invoice.AccountingReference == null ? null : new XElement(Ram + "ReceivableSpecifiedTradeAccountingAccount", Text(Ram + "ID", invoice.AccountingReference)));

    private static XElement? CiiReference(string name, string? reference, string? type = null) => reference == null ? null :
        new XElement(Ram + name, Text(Ram + "IssuerAssignedID", reference), Text(Ram + "TypeCode", type));
    private static XElement? CiiObjectReference(InvoiceIdentifier? identifier) => identifier == null ? null :
        new XElement(Ram + "AdditionalReferencedDocument", Text(Ram + "IssuerAssignedID", identifier.Value), new XElement(Ram + "TypeCode", "130"), Text(Ram + "ReferenceTypeCode", identifier.SchemeId));
    private static XElement? CiiPeriod(InvoicePeriod? period) => period == null ? null :
        new XElement(Ram + "BillingSpecifiedPeriod", CiiDate("StartDateTime", period.Start), CiiDate("EndDateTime", period.End));

    private static XElement CiiTax(InvoiceCalculatedTax tax, Invoice invoice, bool includeTaxPoint) => new XElement(Ram + "ApplicableTradeTax",
        CiiAmount("CalculatedAmount", tax.TaxAmount), new XElement(Ram + "TypeCode", "VAT"), Text(Ram + "ExemptionReason", tax.ExemptionReason),
        CiiAmount("BasisAmount", tax.TaxableAmount), new XElement(Ram + "CategoryCode", tax.CategoryCode), Text(Ram + "ExemptionReasonCode", tax.ExemptionReasonCode),
        includeTaxPoint && invoice.TaxPointDate.HasValue ? new XElement(Ram + "TaxPointDate", new XElement(Udt + "DateString", new XAttribute("format", "102"), invoice.TaxPointDate.Value.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture))) : null,
        includeTaxPoint ? Text(Ram + "DueDateTypeCode", invoice.TaxPointDateCode) : null, tax.Rate.HasValue ? Text(Ram + "RateApplicablePercent", Number(tax.Rate.Value)) : null);

    private static XElement CiiAdjustment(InvoiceAllowanceCharge item, bool documentLevel) => new XElement(Ram + "SpecifiedTradeAllowanceCharge",
        new XElement(Ram + "ChargeIndicator", new XElement(Udt + "Indicator", item.IsCharge ? "true" : "false")),
        item.Percentage.HasValue ? Text(Ram + "CalculationPercent", Number(item.Percentage.Value)) : null,
        item.BaseAmount.HasValue ? CiiAmount("BasisAmount", item.BaseAmount.Value) : null, CiiAmount("ActualAmount", item.Amount),
        Text(Ram + "ReasonCode", item.ReasonCode), Text(Ram + "Reason", item.Reason),
        documentLevel ? new XElement(Ram + "CategoryTradeTax", new XElement(Ram + "TypeCode", "VAT"), Text(Ram + "CategoryCode", item.Tax!.Code),
            item.Tax.Rate.HasValue ? Text(Ram + "RateApplicablePercent", Number(item.Tax.Rate.Value)) : null) : null);
}
