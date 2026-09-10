using OfficeIMO.Invoicing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfInvoiceDocument {
    private static void Paragraph(PdfContentBuilder content, string text) { if (!string.IsNullOrWhiteSpace(text)) content.Paragraph(paragraph => paragraph.Text(text)); }
    private static void DetailGroup(PdfContentBuilder content, string title, string text) =>
        content.Flow(group => { group.H2(title); Paragraph(group, text); }, new PdfFlowOptions { KeepTogether = true });
    private void ComposeDetails(PdfContentBuilder content) {
        var references = new List<string[]>();
        void Add(string label, string? value) { if (!string.IsNullOrWhiteSpace(value)) references.Add(new[] { label, value! }); }
        Add("Purchase order", _invoice.PurchaseOrderReference); Add("Sales order", _invoice.SalesOrderReference);
        Add("Contract", _invoice.ContractReference); Add("Project", _invoice.ProjectReference);
        Add("Despatch advice", _invoice.DespatchAdviceReference); Add("Receiving advice", _invoice.ReceivingAdviceReference);
        Add("Tender", _invoice.TenderReference); Add("Accounting reference", _invoice.AccountingReference);
        Add("Period", _invoice.Period == null ? null : Date(_invoice.Period.Start) + " to " + Date(_invoice.Period.End));
        Add("Tax point", _invoice.TaxPointDate.HasValue ? Date(_invoice.TaxPointDate) : _invoice.TaxPointDateCode);
        foreach (InvoiceReference preceding in _invoice.PrecedingInvoices) Add("Preceding invoice", preceding.Number + " " + Date(preceding.IssueDate));
        if (_invoice.TaxCurrency != null) Add("VAT in accounting currency", _invoice.TaxAmountInAccountingCurrency?.ToString("0.00", FormatCulture) + " " + _invoice.TaxCurrency);
        if (references.Count != 0) { content.H2("References"); content.Table(references, style: PlainTable()); }
        if (_invoice.Delivery != null) {
            DetailGroup(content, "Delivery", Join(_invoice.Delivery.Name, Date(_invoice.Delivery.Date), _invoice.Delivery.LocationIdentifier?.Value,
                _invoice.Delivery.Address == null ? null : Address(_invoice.Delivery.Address)));
        }
        if (_invoice.Payee != null) DetailGroup(content, "Payee", Party(_invoice.Payee));
        if (_invoice.TaxRepresentative != null) DetailGroup(content, "Tax representative", Party(_invoice.TaxRepresentative));
        InvoicePayment? payment = _invoice.Payment;
        if (payment != null || _invoice.PaymentTerms != null) {
            content.H2("Payment");
            if (payment != null) {
                var rows = new List<string[]> { new[] { "Payment means", Join(payment.MeansCode, payment.MeansText).Replace("\n", " ") } };
                if (payment.Reference != null) rows.Add(new[] { "Reference", payment.Reference });
                foreach (InvoiceBankAccount account in payment.Accounts) rows.Add(new[] { account.IsIban ? "IBAN" : "Account", Join(account.Identifier, account.Name, account.ProviderIdentifier) });
                if (payment.MandateReference != null) rows.Add(new[] { "Mandate", payment.MandateReference });
                if (payment.CreditorIdentifier != null) rows.Add(new[] { "Creditor", payment.CreditorIdentifier });
                if (payment.DebitedAccount != null) rows.Add(new[] { "Debited account", payment.DebitedAccount });
                if (payment.CardNumber != null) rows.Add(new[] { "Card", Join("Ending " + payment.CardNumber, payment.CardHolder) });
                content.Table(rows, style: PlainTable());
            }
            if (_invoice.PaymentTerms != null) Paragraph(content, _invoice.PaymentTerms);
        }
        if (_invoice.Notes.Count != 0) {
            content.H2("Notes");
            foreach (InvoiceNote note in _invoice.Notes) Paragraph(content, note.SubjectCode == null ? note.Text : note.SubjectCode + ": " + note.Text);
        }
        if (_invoice.SupportingDocuments.Count != 0) {
            content.H2("Supporting documents");
            content.Table(_invoice.SupportingDocuments.Select(document => new[] {
                document.Reference, Join(document.Description, document.FileName, document.ExternalUri)
            }), style: PlainTable());
        }
    }
}
