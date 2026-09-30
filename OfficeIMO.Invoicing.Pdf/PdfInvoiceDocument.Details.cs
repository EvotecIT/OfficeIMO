using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class PdfInvoiceDocument {
    private void Paragraph(PdfContentBuilder content, string text) { if (!string.IsNullOrWhiteSpace(text)) content.Paragraph(paragraph => paragraph.Text(text), style: _layout.CompactDetails ? new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 4 } : null); }
    private void DetailGroup(PdfContentBuilder content, string title, string text, InvoicePdfTheme? theme = null) =>
        content.Flow(group => { DetailHeading(group, title, theme); Paragraph(group, text); }, new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.MoveToNextPage });
    private void TableGroup(PdfContentBuilder content, string title, IEnumerable<string[]> rows, InvoicePdfTheme? theme = null) =>
        content.Flow(group => { DetailHeading(group, title, theme); group.Table(rows, style: DetailTableStyle(theme)); }, new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.MoveToNextPage });
    private void DetailHeading(PdfContentBuilder content, string title, InvoicePdfTheme? theme) {
        if (_layout.CompactDetails) content.Paragraph(paragraph => paragraph.Bold(title, theme?.Text), style: new PdfParagraphStyle { SpacingBefore = 4, SpacingAfter = 4 });
        else content.H2(title, PdfAlign.Left, theme?.Text);
    }
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
        if (_invoice.Payments.Count != 0 || _invoice.PaymentTerms != null || _invoice.PaymentReference != null ||
            _invoice.CreditorIdentifier != null || _invoice.DirectDebitMandateReference != null) {
            DetailHeading(content, Label(InvoicePdfText.Payment), theme);
            var independent = new List<string[]>();
            if (_invoice.PaymentReference != null) independent.Add(new[] { Label(InvoicePdfText.Reference), _invoice.PaymentReference });
            if (_invoice.CreditorIdentifier != null) independent.Add(new[] { Label(InvoicePdfText.Creditor), _invoice.CreditorIdentifier });
            if (_invoice.DirectDebitMandateReference != null) independent.Add(new[] { Label(InvoicePdfText.Mandate), _invoice.DirectDebitMandateReference });
            if (independent.Count != 0) content.Table(independent, style: DetailTableStyle(theme));
            foreach (InvoicePayment payment in _invoice.Payments) {
                var rows = new List<string[]> { new[] { Label(InvoicePdfText.PaymentMeans), PaymentMeans(payment) } };
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
            DetailHeading(content, Label(InvoicePdfText.Notes), theme);
            foreach (InvoiceNote note in _invoice.Notes) Paragraph(content, note.SubjectCode == null ? note.Text : note.SubjectCode + ": " + note.Text);
        }
        if (_invoice.SupportingDocuments.Count != 0) {
            TableGroup(content, Label(InvoicePdfText.SupportingDocuments), _invoice.SupportingDocuments.Select(document => new[] {
                document.Reference, Join(document.Description, document.FileName, document.MimeType, document.ExternalUri)
            }), theme);
        }
    }

    private PdfTableStyle DetailTableStyle(InvoicePdfTheme? theme) {
        PdfTableStyle style = theme == null ? PlainTable() : new PdfTableStyle {
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
        if (_layout.CompactDetails) { style.SpacingAfter = 4; style.CellPaddingY = 2; }
        return style;
    }
}
