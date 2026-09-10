using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceParser {
    private static Invoice ReadCii(XElement root, InvoiceXmlReadContext c) {
        XElement? context = c.Child(root, Rsm + "ExchangedDocumentContext");
        c.Text(c.Child(context, Ram + "GuidelineSpecifiedDocumentContextParameter"), Ram + "ID");
        XElement? document = c.Child(root, Rsm + "ExchangedDocument");
        XElement? transaction = c.Child(root, Rsm + "SupplyChainTradeTransaction");
        XElement? agreement = c.Child(transaction, Ram + "ApplicableHeaderTradeAgreement");
        XElement? delivery = c.Child(transaction, Ram + "ApplicableHeaderTradeDelivery");
        XElement? settlement = c.Child(transaction, Ram + "ApplicableHeaderTradeSettlement");
        var invoice = new Invoice {
            Number = c.Required(document, Ram + "ID"), TypeCode = c.Required(document, Ram + "TypeCode"), IssueDate = CiiDate(c, document, "IssueDateTime") ?? default,
            BusinessProcessId = c.Text(c.Child(context, Ram + "BusinessProcessSpecifiedDocumentContextParameter"), Ram + "ID"),
            Currency = c.Required(settlement, Ram + "InvoiceCurrencyCode"), TaxCurrency = c.Text(settlement, Ram + "TaxCurrencyCode"),
            BuyerReference = c.Text(agreement, Ram + "BuyerReference"), Seller = CiiParty(c, c.Child(agreement, Ram + "SellerTradeParty")), Buyer = CiiParty(c, c.Child(agreement, Ram + "BuyerTradeParty")),
            PurchaseOrderReference = c.Text(c.Child(agreement, Ram + "BuyerOrderReferencedDocument"), Ram + "IssuerAssignedID"),
            SalesOrderReference = c.Text(c.Child(agreement, Ram + "SellerOrderReferencedDocument"), Ram + "IssuerAssignedID"),
            ContractReference = c.Text(c.Child(agreement, Ram + "ContractReferencedDocument"), Ram + "IssuerAssignedID"),
            DespatchAdviceReference = c.Text(c.Child(delivery, Ram + "DespatchAdviceReferencedDocument"), Ram + "IssuerAssignedID"),
            ReceivingAdviceReference = c.Text(c.Child(delivery, Ram + "ReceivingAdviceReferencedDocument"), Ram + "IssuerAssignedID"),
            AccountingReference = c.Text(c.Child(settlement, Ram + "ReceivableSpecifiedTradeAccountingAccount"), Ram + "ID"), Period = CiiPeriod(c, settlement)
        };
        XElement? project = c.Child(agreement, Ram + "SpecifiedProcuringProject");
        invoice.ProjectReference = c.Text(project, Ram + "ID");
        string? projectName = c.Text(project, Ram + "Name");
        if (projectName != null && projectName != invoice.ProjectReference && projectName != "Project reference")
            c.Loss(project!, "A distinct project name is outside the supported project-reference mapping.");
        XElement? representative = c.Child(agreement, Ram + "SellerTaxRepresentativeTradeParty");
        if (representative != null) invoice.TaxRepresentative = CiiParty(c, representative);
        XElement? payee = c.Child(settlement, Ram + "PayeeTradeParty");
        if (payee != null) invoice.Payee = CiiParty(c, payee);
        foreach (XElement note in c.Children(document, Ram + "IncludedNote")) invoice.Notes.Add(new InvoiceNote(c.Required(note, Ram + "Content"), c.Text(note, Ram + "SubjectCode")));
        foreach (XElement reference in c.Children(agreement, Ram + "AdditionalReferencedDocument")) CiiAdditionalReference(c, reference, invoice);
        XElement? shipTo = c.Child(delivery, Ram + "ShipToTradeParty");
        XElement? deliveryEvent = c.Child(delivery, Ram + "ActualDeliverySupplyChainEvent");
        if (shipTo != null || deliveryEvent != null) {
            XElement? address = c.Child(shipTo, Ram + "PostalTradeAddress");
            InvoiceIdentifier? localId = c.Identifier(c.Child(shipTo, Ram + "ID"));
            InvoiceIdentifier? globalId = c.Identifier(c.Child(shipTo, Ram + "GlobalID"));
            if (localId != null && globalId != null) c.Loss(shipTo!, "Multiple delivery location identifiers are outside the supported mapping.");
            invoice.Delivery = new InvoiceDelivery { Name = c.Text(shipTo, Ram + "Name"), LocationIdentifier = globalId ?? localId,
                Address = address == null ? null : CiiAddress(c, address), Date = CiiDate(c, deliveryEvent, "OccurrenceDateTime") };
        }
        foreach (XElement line in c.Children(transaction, Ram + "IncludedSupplyChainTradeLineItem")) {
            if (invoice.Lines.Count >= 10000) throw new InvalidDataException("Invoice exceeds 10,000 lines.");
            invoice.Lines.Add(CiiLine(c, line, invoice.Currency));
        }
        invoice.Payment = CiiPayment(c, settlement);
        XElement? terms = c.Child(settlement, Ram + "SpecifiedTradePaymentTerms");
        invoice.PaymentTerms = c.Text(terms, Ram + "Description"); invoice.DueDate = CiiDate(c, terms, "DueDateDateTime");
        string? mandate = c.Text(terms, Ram + "DirectDebitMandateID");
        if (mandate != null) { invoice.Payment = invoice.Payment ?? new InvoicePayment(); invoice.Payment.MandateReference = mandate; }
        foreach (XElement adjustment in c.Children(settlement, Ram + "SpecifiedTradeAllowanceCharge")) invoice.AllowancesAndCharges.Add(CiiAdjustment(c, adjustment, invoice.Currency, true));
        foreach (XElement reference in c.Children(settlement, Ram + "InvoiceReferencedDocument"))
            invoice.PrecedingInvoices.Add(new InvoiceReference(c.Required(reference, Ram + "IssuerAssignedID"), CiiDate(c, reference, "FormattedIssueDateTime", true)));
        foreach (XElement tax in c.Children(settlement, Ram + "ApplicableTradeTax")) {
            var declared = new InvoiceDeclaredTax { Category = CiiTaxCategory(c, tax), TaxableAmount = c.RequiredMoney(tax, Ram + "BasisAmount", invoice.Currency),
                TaxAmount = c.RequiredMoney(tax, Ram + "CalculatedAmount", invoice.Currency) };
            XElement? point = c.Child(tax, Ram + "TaxPointDate");
            DateTime? date = c.Date(c.Child(point, Udt + "DateString"), true) ?? c.Date(c.Child(point, Udt + "Date"));
            string? code = c.Text(tax, Ram + "DueDateTypeCode");
            if (date.HasValue) {
                if (invoice.TaxPointDate.HasValue) c.Loss(tax, "Invoice-wide tax point date is declared more than once.");
                invoice.TaxPointDate = invoice.TaxPointDate ?? date;
            }
            if (code != null) {
                if (invoice.TaxPointDateCode != null) c.Loss(tax, "Invoice-wide tax point code is declared more than once.");
                invoice.TaxPointDateCode = invoice.TaxPointDateCode ?? code;
            }
            invoice.DeclaredTaxes.Add(declared);
        }
        CiiTotals(c, c.Child(settlement, Ram + "SpecifiedTradeSettlementHeaderMonetarySummation"), invoice);
        return invoice;
    }

    private static void CiiAdditionalReference(InvoiceXmlReadContext c, XElement reference, Invoice invoice) {
        string? type = c.Text(reference, Ram + "TypeCode");
        string number = c.Required(reference, Ram + "IssuerAssignedID");
        if (type == "50") {
            if (invoice.TenderReference != null) c.Loss(reference, "Multiple tender references are outside the supported mapping.");
            invoice.TenderReference = number; return;
        }
        if (type == "130") {
            if (invoice.ObjectIdentifier != null) c.Loss(reference, "Multiple invoiced-object references are outside the supported mapping.");
            invoice.ObjectIdentifier = new InvoiceIdentifier(number, c.Text(reference, Ram + "ReferenceTypeCode")); return;
        }
        if (type != "916") c.Loss(reference, "Supporting document type is not 916.");
        XElement? binary = c.Child(reference, Ram + "AttachmentBinaryObject");
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument { Reference = number, Description = c.Text(reference, Ram + "Name"), ExternalUri = c.Text(reference, Ram + "URIID"),
            Data = Binary(c, binary), FileName = c.Attribute(binary, "filename"), MimeType = c.Attribute(binary, "mimeCode") });
    }

    private static InvoiceTaxCategory CiiTaxCategory(InvoiceXmlReadContext c, XElement? element) {
        c.Expected(c.Require(element, Ram + "TypeCode"), "VAT");
        return new InvoiceTaxCategory { Code = c.Required(element, Ram + "CategoryCode"), Rate = c.Decimal(element, Ram + "RateApplicablePercent"),
            ExemptionReason = c.Text(element, Ram + "ExemptionReason"), ExemptionReasonCode = c.Text(element, Ram + "ExemptionReasonCode") };
    }
    private static InvoiceAllowanceCharge CiiAdjustment(InvoiceXmlReadContext c, XElement element, string currency, bool documentLevel) => new InvoiceAllowanceCharge {
        IsCharge = c.Boolean(c.Child(c.Child(element, Ram + "ChargeIndicator"), Udt + "Indicator")), Amount = c.RequiredMoney(element, Ram + "ActualAmount", currency),
        BaseAmount = c.Money(element, Ram + "BasisAmount", currency), Percentage = c.Decimal(element, Ram + "CalculationPercent"),
        Reason = c.Text(element, Ram + "Reason"), ReasonCode = c.Text(element, Ram + "ReasonCode"),
        Tax = documentLevel ? CiiTaxCategory(c, c.Child(element, Ram + "CategoryTradeTax")) : null
    };
    private static void CiiTotals(InvoiceXmlReadContext c, XElement? totals, Invoice invoice) {
        if (totals == null) throw new InvalidDataException("CII invoice monetary totals are required.");
        invoice.DeclaredTotals = new InvoiceDeclaredTotals {
            LineNetTotal = c.RequiredMoney(totals, Ram + "LineTotalAmount", invoice.Currency), AllowanceTotal = c.Money(totals, Ram + "AllowanceTotalAmount", invoice.Currency),
            ChargeTotal = c.Money(totals, Ram + "ChargeTotalAmount", invoice.Currency), TaxExclusiveTotal = c.RequiredMoney(totals, Ram + "TaxBasisTotalAmount", invoice.Currency),
            TaxInclusiveTotal = c.RequiredMoney(totals, Ram + "GrandTotalAmount", invoice.Currency), PayableAmount = c.RequiredMoney(totals, Ram + "DuePayableAmount", invoice.Currency)
        };
        invoice.PrepaidAmount = c.Money(totals, Ram + "TotalPrepaidAmount", invoice.Currency) ?? 0m;
        invoice.RoundingAmount = c.Money(totals, Ram + "RoundingAmount", invoice.Currency) ?? 0m;
        foreach (XElement tax in c.Children(totals, Ram + "TaxTotalAmount")) {
            string? currency = c.Attribute(tax, "currencyID");
            decimal? value = c.Decimal(tax);
            if (currency == null || currency == invoice.Currency) {
                if (invoice.DeclaredTotals.TaxTotal.HasValue) c.Loss(tax, "Invoice currency VAT total is duplicated.");
                invoice.DeclaredTotals.TaxTotal = value;
            } else if (currency == invoice.TaxCurrency) {
                if (invoice.TaxAmountInAccountingCurrency.HasValue) c.Loss(tax, "Accounting currency VAT total is duplicated.");
                invoice.TaxAmountInAccountingCurrency = value;
            } else c.Loss(tax, "VAT amount uses an undeclared currency.");
        }
        if (!invoice.DeclaredTotals.TaxTotal.HasValue || invoice.DeclaredTaxes.Count == 0) throw new InvalidDataException("CII VAT total and breakdown are required.");
    }
}
