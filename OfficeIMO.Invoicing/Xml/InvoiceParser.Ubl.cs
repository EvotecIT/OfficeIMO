using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceParser {
    private static Invoice ReadUbl(XElement root, InvoiceXmlReadContext c) {
        bool credit = root.Name.Namespace == InvoiceXml.UblCreditNote;
        c.Expected(c.Child(root, Cbc + "UBLVersionID"), "2.1");
        c.Text(root, Cbc + "CustomizationID");
        var invoice = new Invoice {
            Number = c.Required(root, Cbc + "ID"), IssueDate = c.Date(c.Child(root, Cbc + "IssueDate")) ?? default,
            TypeCode = c.Required(root, Cbc + (credit ? "CreditNoteTypeCode" : "InvoiceTypeCode")), Currency = c.Required(root, Cbc + "DocumentCurrencyCode"),
            TaxCurrency = c.Text(root, Cbc + "TaxCurrencyCode"), BusinessProcessId = c.Text(root, Cbc + "ProfileID"), BuyerReference = c.Text(root, Cbc + "BuyerReference"),
            DueDate = c.Date(c.Child(root, Cbc + "DueDate")), TaxPointDate = c.Date(c.Child(root, Cbc + "TaxPointDate")), AccountingReference = c.Text(root, Cbc + "AccountingCost"),
            ContractReference = c.Text(c.Child(root, Cac + "ContractDocumentReference"), Cbc + "ID"),
            DespatchAdviceReference = c.Text(c.Child(root, Cac + "DespatchDocumentReference"), Cbc + "ID"),
            ReceivingAdviceReference = c.Text(c.Child(root, Cac + "ReceiptDocumentReference"), Cbc + "ID"),
            TenderReference = c.Text(c.Child(root, Cac + "OriginatorDocumentReference"), Cbc + "ID"),
            ProjectReference = c.Text(c.Child(root, Cac + "ProjectReference"), Cbc + "ID")
        };
        if (credit && invoice.TypeCode != "381") c.Loss(root, "Credit note type is outside the supported UBL credit note mapping.");
        if (!credit && invoice.TypeCode == "381") c.Loss(root, "An Invoice root cannot be converted as a CreditNote merely because its type code is 381.");
        invoice.Period = UblPeriod(c, root, out string? taxPointCode); invoice.TaxPointDateCode = taxPointCode;
        foreach (XElement note in c.Children(root, Cbc + "Note")) {
            string text = c.Value(note)!;
            if (InvoiceNote.HasEncodedSubject(text))
                invoice.Notes.Add(new InvoiceNote(text.Substring(5), text.Substring(1, 3)));
            else invoice.Notes.Add(new InvoiceNote(text));
        }
        XElement? order = c.Child(root, Cac + "OrderReference");
        invoice.PurchaseOrderReference = c.Text(order, Cbc + "ID"); invoice.SalesOrderReference = c.Text(order, Cbc + "SalesOrderID");
        foreach (XElement billing in c.Children(root, Cac + "BillingReference")) {
            XElement? reference = c.Child(billing, Cac + "InvoiceDocumentReference");
            invoice.PrecedingInvoices.Add(new InvoiceReference(c.Required(reference, Cbc + "ID"), c.Date(c.Child(reference, Cbc + "IssueDate"))));
        }
        foreach (XElement reference in c.Children(root, Cac + "AdditionalDocumentReference")) UblAdditionalReference(c, reference, invoice);
        invoice.Seller = UblParty(c, c.Child(c.Child(root, Cac + "AccountingSupplierParty"), Cac + "Party"), out string? creditor);
        invoice.Buyer = UblParty(c, c.Child(c.Child(root, Cac + "AccountingCustomerParty"), Cac + "Party"), out string? buyerCreditor);
        if (buyerCreditor != null) c.Loss(root, "SEPA creditor identifier belongs to the seller, not the buyer.");
        XElement? payee = c.Child(root, Cac + "PayeeParty");
        if (payee != null) {
            invoice.Payee = new InvoiceParty { Name = c.Required(c.Child(payee, Cac + "PartyName"), Cbc + "Name"),
                LegalRegistration = c.Identifier(c.Child(c.Child(payee, Cac + "PartyLegalEntity"), Cbc + "CompanyID")) };
            foreach (XElement identifier in c.Children(payee, Cac + "PartyIdentification")) invoice.Payee.Identifiers.Add(c.Identifier(c.Child(identifier, Cbc + "ID")) ?? new InvoiceIdentifier(string.Empty));
        }
        XElement? representative = c.Child(root, Cac + "TaxRepresentativeParty");
        if (representative != null) {
            invoice.TaxRepresentative = new InvoiceParty { Name = c.Required(c.Child(representative, Cac + "PartyName"), Cbc + "Name"), Address = UblAddress(c, c.Child(representative, Cac + "PostalAddress")) };
            UblTaxRegistrations(c, representative, invoice.TaxRepresentative);
        }
        XElement? delivery = c.Child(root, Cac + "Delivery");
        if (delivery != null) {
            XElement? location = c.Child(delivery, Cac + "DeliveryLocation"), address = c.Child(location, Cac + "Address");
            invoice.Delivery = new InvoiceDelivery { Date = c.Date(c.Child(delivery, Cbc + "ActualDeliveryDate")),
                Name = c.Text(c.Child(c.Child(delivery, Cac + "DeliveryParty"), Cac + "PartyName"), Cbc + "Name"),
                LocationIdentifier = c.Identifier(c.Child(location, Cbc + "ID")), Address = address == null ? null : UblAddress(c, address) };
        }
        invoice.Payment = UblPayment(c, root, creditor);
        invoice.PaymentTerms = c.Text(c.Child(root, Cac + "PaymentTerms"), Cbc + "Note");
        foreach (XElement adjustment in c.Children(root, Cac + "AllowanceCharge")) invoice.AllowancesAndCharges.Add(UblAdjustment(c, adjustment, invoice.Currency, true));
        foreach (XElement line in c.Children(root, Cac + (credit ? "CreditNoteLine" : "InvoiceLine"))) {
            if (invoice.Lines.Count >= 10000) throw new InvalidDataException("Invoice exceeds 10,000 lines.");
            invoice.Lines.Add(UblLine(c, line, invoice.Currency, credit));
        }
        UblTotals(c, root, invoice);
        return invoice;
    }

    private static void UblAdditionalReference(InvoiceXmlReadContext c, XElement reference, Invoice invoice) {
        XElement? identifier = c.Child(reference, Cbc + "ID");
        string? type = c.Text(reference, Cbc + "DocumentTypeCode");
        if (type == "130") {
            if (invoice.ObjectIdentifier != null) c.Loss(reference, "Multiple invoiced-object references are outside the supported mapping.");
            invoice.ObjectIdentifier = c.Identifier(identifier); return;
        }
        if (type != null) c.Loss(reference, "Supporting document type is outside the supported UBL mapping.");
        XElement? attachment = c.Child(reference, Cac + "Attachment"), binary = c.Child(attachment, Cbc + "EmbeddedDocumentBinaryObject");
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument { Reference = c.Value(identifier) ?? string.Empty, Description = c.Text(reference, Cbc + "DocumentDescription"),
            ExternalUri = c.Text(c.Child(attachment, Cac + "ExternalReference"), Cbc + "URI"), Data = Binary(c, binary), FileName = c.Attribute(binary, "filename"), MimeType = c.Attribute(binary, "mimeCode") });
    }
    private static InvoiceTaxCategory UblTaxCategory(InvoiceXmlReadContext c, XElement? element) {
        c.Expected(c.Require(c.Child(element, Cac + "TaxScheme"), Cbc + "ID"), "VAT");
        return new InvoiceTaxCategory { Code = c.Required(element, Cbc + "ID"), Rate = c.Decimal(element, Cbc + "Percent"),
            ExemptionReason = c.Text(element, Cbc + "TaxExemptionReason"), ExemptionReasonCode = c.Text(element, Cbc + "TaxExemptionReasonCode") };
    }
    private static InvoiceAllowanceCharge UblAdjustment(InvoiceXmlReadContext c, XElement element, string currency, bool documentLevel) => new InvoiceAllowanceCharge {
        IsCharge = c.Boolean(c.Child(element, Cbc + "ChargeIndicator")), Amount = c.RequiredMoney(element, Cbc + "Amount", currency, true),
        BaseAmount = c.Money(element, Cbc + "BaseAmount", currency, true), Percentage = c.Decimal(element, Cbc + "MultiplierFactorNumeric"),
        Reason = c.Text(element, Cbc + "AllowanceChargeReason"), ReasonCode = c.Text(element, Cbc + "AllowanceChargeReasonCode"),
        Tax = documentLevel ? UblTaxCategory(c, c.Child(element, Cac + "TaxCategory")) : null
    };
    private static void UblTotals(InvoiceXmlReadContext c, XElement root, Invoice invoice) {
        XElement? totals = c.Child(root, Cac + "LegalMonetaryTotal");
        if (totals == null) throw new InvalidDataException("UBL invoice monetary totals are required.");
        invoice.DeclaredTotals = new InvoiceDeclaredTotals {
            LineNetTotal = c.RequiredMoney(totals, Cbc + "LineExtensionAmount", invoice.Currency, true), AllowanceTotal = c.Money(totals, Cbc + "AllowanceTotalAmount", invoice.Currency, true),
            ChargeTotal = c.Money(totals, Cbc + "ChargeTotalAmount", invoice.Currency, true), TaxExclusiveTotal = c.RequiredMoney(totals, Cbc + "TaxExclusiveAmount", invoice.Currency, true),
            TaxInclusiveTotal = c.RequiredMoney(totals, Cbc + "TaxInclusiveAmount", invoice.Currency, true), PayableAmount = c.RequiredMoney(totals, Cbc + "PayableAmount", invoice.Currency, true)
        };
        invoice.PrepaidAmount = c.Money(totals, Cbc + "PrepaidAmount", invoice.Currency, true) ?? 0m;
        invoice.RoundingAmount = c.Money(totals, Cbc + "PayableRoundingAmount", invoice.Currency, true) ?? 0m;
        foreach (XElement tax in c.Children(root, Cac + "TaxTotal")) {
            XElement? amount = c.Child(tax, Cbc + "TaxAmount");
            string? currency = c.Attribute(amount, "currencyID");
            decimal? value = c.Decimal(amount);
            if (currency == invoice.Currency) {
                invoice.DeclaredTotals = invoice.DeclaredTotals ?? new InvoiceDeclaredTotals();
                if (invoice.DeclaredTotals.TaxTotal.HasValue) c.Loss(tax, "Invoice currency VAT total is duplicated.");
                invoice.DeclaredTotals.TaxTotal = value;
            } else if (currency != null && currency == invoice.TaxCurrency) {
                if (invoice.TaxAmountInAccountingCurrency.HasValue) c.Loss(tax, "Accounting currency VAT total is duplicated.");
                invoice.TaxAmountInAccountingCurrency = value;
            } else c.Loss(tax, "VAT total currency is missing or undeclared.");
            foreach (XElement subtotal in c.Children(tax, Cac + "TaxSubtotal")) {
                if (currency != invoice.Currency) c.Loss(subtotal, "Accounting-currency VAT breakdowns cannot be mapped to invoice-currency totals.");
                invoice.DeclaredTaxes.Add(new InvoiceDeclaredTax { Category = UblTaxCategory(c, c.Child(subtotal, Cac + "TaxCategory")),
                    TaxableAmount = c.RequiredMoney(subtotal, Cbc + "TaxableAmount", invoice.Currency, true), TaxAmount = c.RequiredMoney(subtotal, Cbc + "TaxAmount", invoice.Currency, true) });
            }
        }
        if (!invoice.DeclaredTotals.TaxTotal.HasValue || invoice.DeclaredTaxes.Count == 0) throw new InvalidDataException("UBL VAT total and breakdown are required.");
    }
}
