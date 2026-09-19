using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class PdfInvoiceDocument {
    private static void Paragraph(PdfContentBuilder content, string text) { if (!string.IsNullOrWhiteSpace(text)) content.Paragraph(paragraph => paragraph.Text(text)); }
    private static void DetailGroup(PdfContentBuilder content, string title, string text, InvoicePdfTheme? theme = null) =>
        content.Flow(group => { group.H2(title, PdfAlign.Left, theme?.Text); Paragraph(group, text); }, new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.MoveToNextPage });
    private static void TableGroup(PdfContentBuilder content, string title, IEnumerable<string[]> rows, InvoicePdfTheme? theme = null) =>
        content.Flow(group => { group.H2(title, PdfAlign.Left, theme?.Text); group.Table(rows, style: DetailTableStyle(theme)); }, new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.MoveToNextPage });
    private void ComposeDetails(PdfContentBuilder content, InvoicePdfTheme? theme = null) {
        var references = new List<string[]>();
        void Add(string label, string? value) { if (!string.IsNullOrWhiteSpace(value)) references.Add(new[] { label, value! }); }
        Add(Label(InvoicePdfText.PurchaseOrder), _invoice.PurchaseOrderReference); Add(Label(InvoicePdfText.SalesOrder), _invoice.SalesOrderReference);
        Add(Label(InvoicePdfText.Contract), _invoice.ContractReference); Add(Label(InvoicePdfText.Project), _invoice.ProjectReference);
        Add(Label(InvoicePdfText.DespatchAdvice), _invoice.DespatchAdviceReference); Add(Label(InvoicePdfText.ReceivingAdvice), _invoice.ReceivingAdviceReference);
        Add(Label(InvoicePdfText.Tender), _invoice.TenderReference); Add(Label(InvoicePdfText.AccountingReference), _invoice.AccountingReference);
        Add(Label(InvoicePdfText.InvoicedObject), Identifier(Label(InvoicePdfText.Identifier), _invoice.ObjectIdentifier));
        Add(Label(InvoicePdfText.BusinessProcess), _invoice.BusinessProcessId);
        Add(Label(InvoicePdfText.DocumentType), _invoice.TypeCode);
        Add(Label(InvoicePdfText.Period), Period(_invoice.Period));
        Add(Label(InvoicePdfText.TaxPoint), _invoice.TaxPointDate.HasValue ? Date(_invoice.TaxPointDate) : _invoice.TaxPointDateCode);
        foreach (InvoiceReference preceding in _invoice.PrecedingInvoices) Add(Label(InvoicePdfText.PrecedingInvoice), preceding.Number + " " + Date(preceding.IssueDate));
        if (_invoice.TaxCurrency != null) Add(Label(InvoicePdfText.VatAccountingCurrency), _invoice.TaxAmountInAccountingCurrency?.ToString("0.00", _layout.FormattingCulture) + " " + _invoice.TaxCurrency);
        if (references.Count != 0) TableGroup(content, Label(InvoicePdfText.References), references, theme);
        if (_invoice.Delivery != null) {
            DetailGroup(content, Label(InvoicePdfText.Delivery), Join(_invoice.Delivery.Name, Date(_invoice.Delivery.Date), Identifier(Label(InvoicePdfText.Location), _invoice.Delivery.LocationIdentifier),
                _invoice.Delivery.Address == null ? null : Address(_invoice.Delivery.Address)), theme);
        }
        if (_invoice.Payee != null) DetailGroup(content, Label(InvoicePdfText.Payee), Party(_invoice.Payee), theme);
        if (_invoice.TaxRepresentative != null) DetailGroup(content, Label(InvoicePdfText.TaxRepresentative), Party(_invoice.TaxRepresentative), theme);
        if (_invoice.Payments.Count != 0 || _invoice.PaymentTerms != null) {
            content.H2(Label(InvoicePdfText.Payment), PdfAlign.Left, theme?.Text);
            foreach (InvoicePayment payment in _invoice.Payments) {
                var rows = new List<string[]> { new[] { Label(InvoicePdfText.PaymentMeans), Join(payment.MeansCode, payment.MeansText).Replace("\n", " ") } };
                if (payment.Reference != null) rows.Add(new[] { Label(InvoicePdfText.Reference), payment.Reference });
                if (payment.Account != null) rows.Add(new[] { payment.Account.IsIban ? "IBAN" : Label(InvoicePdfText.Account), Join(payment.Account.Identifier, payment.Account.Name, payment.Account.ProviderIdentifier) });
                if (payment.MandateReference != null) rows.Add(new[] { Label(InvoicePdfText.Mandate), payment.MandateReference });
                if (payment.CreditorIdentifier != null) rows.Add(new[] { Label(InvoicePdfText.Creditor), payment.CreditorIdentifier });
                if (payment.DebitedAccount != null) rows.Add(new[] { Label(InvoicePdfText.DebitedAccount), payment.DebitedAccount });
                if (payment.CardNumber != null) rows.Add(new[] { Label(InvoicePdfText.Card), Join(Label(InvoicePdfText.Ending) + " " + payment.CardNumber, payment.CardNetworkId, payment.CardHolder) });
                content.Table(rows, style: DetailTableStyle(theme));
            }
            if (_invoice.PaymentTerms != null) Paragraph(content, _invoice.PaymentTerms);
        }
        if (_invoice.Notes.Count != 0) {
            content.H2(Label(InvoicePdfText.Notes), PdfAlign.Left, theme?.Text);
            foreach (InvoiceNote note in _invoice.Notes) Paragraph(content, note.SubjectCode == null ? note.Text : note.SubjectCode + ": " + note.Text);
        }
        if (_invoice.SupportingDocuments.Count != 0) {
            TableGroup(content, Label(InvoicePdfText.SupportingDocuments), _invoice.SupportingDocuments.Select(document => new[] {
                document.Reference, Join(document.Description, document.FileName, document.MimeType, document.ExternalUri)
            }), theme);
        }
    }

    private static PdfTableStyle DetailTableStyle(InvoicePdfTheme? theme) => theme == null
        ? PlainTable()
        : new PdfTableStyle {
            HeaderRowCount = 0,
            FontSize = 9D,
            SpacingAfter = 10D,
            RowStripeFill = null,
            BorderColor = theme.Border,
            BorderWidth = 0.6D,
            CornerRadius = theme.CornerRadius,
            CellPaddingX = 7D,
            CellPaddingY = 4D,
            BodyColumnFills = new List<PdfColor?> { theme.Surface, null },
            ColumnWidthWeights = new List<double> { 1D, 1D }
        };
}
