using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static XDocument WriteUbl(Invoice invoice, InvoiceCalculation calculation, InvoiceXmlOptions options) {
        bool credit = invoice.TypeCode == "381";
        XNamespace rootNamespace = credit ? InvoiceXml.UblCreditNote : InvoiceXml.UblInvoice;
        return new XDocument(new XElement(rootNamespace + (credit ? "CreditNote" : "Invoice"),
            new XAttribute("xmlns", rootNamespace), new XAttribute(XNamespace.Xmlns + "cac", Cac), new XAttribute(XNamespace.Xmlns + "cbc", Cbc),
            Text(Cbc + "CustomizationID", InvoiceProfiles.GetGuidelineId(options.Profile)), Text(Cbc + "ProfileID", invoice.BusinessProcessId),
            Text(Cbc + "ID", invoice.Number), UblDate("IssueDate", invoice.IssueDate), credit ? UblDate("TaxPointDate", invoice.TaxPointDate) : UblDate("DueDate", invoice.DueDate),
            Text(Cbc + (credit ? "CreditNoteTypeCode" : "InvoiceTypeCode"), invoice.TypeCode),
            invoice.Notes.Select(note => Text(Cbc + "Note", note.SubjectCode == null ? note.Text : "#" + note.SubjectCode + "#" + note.Text)),
            credit ? null : UblDate("TaxPointDate", invoice.TaxPointDate), Text(Cbc + "DocumentCurrencyCode", invoice.Currency), Text(Cbc + "TaxCurrencyCode", invoice.TaxCurrency),
            Text(Cbc + "AccountingCost", invoice.AccountingReference), Text(Cbc + "BuyerReference", invoice.BuyerReference), UblPeriod(invoice.Period, invoice.TaxPointDateCode),
            invoice.PurchaseOrderReference == null ? null : new XElement(Cac + "OrderReference", Text(Cbc + "ID", invoice.PurchaseOrderReference), Text(Cbc + "SalesOrderID", invoice.SalesOrderReference)),
            invoice.PrecedingInvoices.Select(reference => new XElement(Cac + "BillingReference", new XElement(Cac + "InvoiceDocumentReference", Text(Cbc + "ID", reference.Number), UblDate("IssueDate", reference.IssueDate)))),
            UblReference("DespatchDocumentReference", invoice.DespatchAdviceReference), UblReference("ReceiptDocumentReference", invoice.ReceivingAdviceReference),
            credit ? null : UblReference("OriginatorDocumentReference", invoice.TenderReference), UblReference("ContractDocumentReference", invoice.ContractReference),
            invoice.SupportingDocuments.Select(UblSupportingDocument), UblObjectReference(invoice.ObjectIdentifier),
            credit ? UblReference("OriginatorDocumentReference", invoice.TenderReference) : UblReference("ProjectReference", invoice.ProjectReference),
            new XElement(Cac + "AccountingSupplierParty", UblParty(invoice.Seller, invoice.Payment?.CreditorIdentifier)),
            new XElement(Cac + "AccountingCustomerParty", UblParty(invoice.Buyer)),
            invoice.Payee == null ? null : new XElement(Cac + "PayeeParty",
                invoice.Payee.Identifiers.Select(identifier => new XElement(Cac + "PartyIdentification", Identifier(Cbc + "ID", identifier))),
                new XElement(Cac + "PartyName", Text(Cbc + "Name", invoice.Payee.Name)),
                invoice.Payee.LegalRegistration == null ? null : new XElement(Cac + "PartyLegalEntity", Identifier(Cbc + "CompanyID", invoice.Payee.LegalRegistration))),
            invoice.TaxRepresentative == null ? null : new XElement(Cac + "TaxRepresentativeParty", new XElement(Cac + "PartyName", Text(Cbc + "Name", invoice.TaxRepresentative.Name)),
                UblAddress("PostalAddress", invoice.TaxRepresentative.Address), UblTaxRegistration(invoice.TaxRepresentative.VatIdentifier, "VAT")),
            UblDelivery(invoice.Delivery), UblPayment(invoice.Payment),
            invoice.PaymentTerms == null ? null : new XElement(Cac + "PaymentTerms", Text(Cbc + "Note", invoice.PaymentTerms)),
            invoice.AllowancesAndCharges.Select(item => UblAdjustment(item, invoice.Currency, true)),
            new XElement(Cac + "TaxTotal", UblAmount("TaxAmount", calculation.TaxTotal, invoice.Currency), calculation.Taxes.Select(tax => new XElement(Cac + "TaxSubtotal",
                UblAmount("TaxableAmount", tax.TaxableAmount, invoice.Currency), UblAmount("TaxAmount", tax.TaxAmount, invoice.Currency),
                UblTaxCategory("TaxCategory", tax.CategoryCode, tax.Rate, tax.ExemptionReason, tax.ExemptionReasonCode)))),
            invoice.TaxAmountInAccountingCurrency.HasValue ? new XElement(Cac + "TaxTotal", UblAmount("TaxAmount", invoice.TaxAmountInAccountingCurrency.Value, invoice.TaxCurrency!)) : null,
            new XElement(Cac + "LegalMonetaryTotal", UblAmount("LineExtensionAmount", calculation.LineNetTotal, invoice.Currency),
                UblAmount("TaxExclusiveAmount", calculation.TaxExclusiveTotal, invoice.Currency), UblAmount("TaxInclusiveAmount", calculation.TaxInclusiveTotal, invoice.Currency),
                UblAmount("AllowanceTotalAmount", calculation.AllowanceTotal, invoice.Currency), UblAmount("ChargeTotalAmount", calculation.ChargeTotal, invoice.Currency),
                UblAmount("PrepaidAmount", calculation.PrepaidAmount, invoice.Currency), UblAmount("PayableRoundingAmount", calculation.RoundingAmount, invoice.Currency),
                UblAmount("PayableAmount", calculation.PayableAmount, invoice.Currency)),
            invoice.Lines.Select((line, index) => UblLine(line, calculation.Lines[index], invoice.Currency, credit))));
    }

    private static XElement? UblReference(string name, string? reference) => reference == null ? null : new XElement(Cac + name, Text(Cbc + "ID", reference));
    private static XElement? UblObjectReference(InvoiceIdentifier? identifier) => identifier == null ? null : new XElement(Cac + "AdditionalDocumentReference",
        Identifier(Cbc + "ID", identifier), new XElement(Cbc + "DocumentTypeCode", "130"));
    private static XElement UblSupportingDocument(InvoiceSupportingDocument document) => new XElement(Cac + "AdditionalDocumentReference",
        Text(Cbc + "ID", document.Reference), Text(Cbc + "DocumentDescription", document.Description),
        document.Data == null && document.ExternalUri == null ? null : new XElement(Cac + "Attachment",
            document.Data == null ? null : new XElement(Cbc + "EmbeddedDocumentBinaryObject", new XAttribute("mimeCode", document.MimeType!), new XAttribute("filename", document.FileName!), Convert.ToBase64String(document.Data)),
            document.ExternalUri == null ? null : new XElement(Cac + "ExternalReference", Text(Cbc + "URI", document.ExternalUri))));
    private static XElement? UblPeriod(InvoicePeriod? period, string? taxPointCode = null) => period == null && taxPointCode == null ? null :
        new XElement(Cac + "InvoicePeriod", UblDate("StartDate", period?.Start), UblDate("EndDate", period?.End), Text(Cbc + "DescriptionCode", taxPointCode));
    private static XElement UblTaxCategory(string name, string code, decimal? rate, string? reason = null, string? reasonCode = null) => new XElement(Cac + name,
        Text(Cbc + "ID", code), rate.HasValue ? Text(Cbc + "Percent", Number(rate.Value)) : null,
        Text(Cbc + "TaxExemptionReasonCode", reasonCode), Text(Cbc + "TaxExemptionReason", reason), new XElement(Cac + "TaxScheme", new XElement(Cbc + "ID", "VAT")));
    private static XElement UblAdjustment(InvoiceAllowanceCharge item, string currency, bool documentLevel) => new XElement(Cac + "AllowanceCharge",
        new XElement(Cbc + "ChargeIndicator", item.IsCharge ? "true" : "false"), Text(Cbc + "AllowanceChargeReasonCode", item.ReasonCode), Text(Cbc + "AllowanceChargeReason", item.Reason),
        item.Percentage.HasValue ? Text(Cbc + "MultiplierFactorNumeric", Number(item.Percentage.Value)) : null,
        UblAmount("Amount", item.Amount, currency), item.BaseAmount.HasValue ? UblAmount("BaseAmount", item.BaseAmount.Value, currency) : null,
        documentLevel ? UblTaxCategory("TaxCategory", item.Tax!.Code, item.Tax.Rate) : null);
}
