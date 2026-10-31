using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class PdfInvoiceDocument {
    private void ComposeModern(PdfContentBuilder content, InvoicePdfTheme theme) {
        content.Row(row => {
            row.Style(new PdfRowStyle { Gap = global::OfficeIMO.OfficeLength.Points(20D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(16D) })
                .PercentColumn(44D, column => {
                    if (_layout.LogoBytesSnapshot != null) {
                        column.Image(
                            _layout.LogoBytesSnapshot,
                            global::OfficeIMO.OfficeLength.Points(_layout.LogoWidth),
                            global::OfficeIMO.OfficeLength.Points(_layout.LogoHeight),
                            align: PdfAlign.Left,
                            clipPath: null,
                            fit: _layout.LogoFit,
                            spacingBefore: null,
                            spacingAfter: null,
                            style: null,
                            linkUri: null,
                            linkContents: null,
                            alternativeText: _layout.LogoAlternativeText);
                    } else {
                        column.H2(_invoice.Seller.Name, PdfAlign.Left, theme.Text, style: SectionHeadingStyle());
                    }
                })
                .PercentColumn(56D, column => {
                    column.Paragraph(p => p.FontSize(global::OfficeIMO.OfficeLength.Points(18D)).Bold(DocumentTitle, theme.Text), PdfAlign.Right, style: new PdfParagraphStyle {
                        SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
                        SpacingAfter = global::OfficeIMO.OfficeLength.Points(6D),
                        WidowControl = true
                    });
                    column.Paragraph(p => p
                        .Color(theme.MutedText)
                        .Text(Label(InvoicePdfText.Issued) + ": ")
                        .Bold(Date(_invoice.IssueDate), theme.Text)
                        .Color(theme.MutedText)
                        .Text("\n" + Label(InvoicePdfText.Due) + ": ")
                        .Bold(Date(_invoice.DueDate), theme.Text), PdfAlign.Right);
                });
        });

        content.Row(row => row
            .Style(new PdfRowStyle { Gap = global::OfficeIMO.OfficeLength.Points(12D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(12D) })
            .PercentColumn(50D, column => PartyCard(column, Label(InvoicePdfText.Seller), _invoice.Seller, theme))
            .PercentColumn(50D, column => PartyCard(column, Label(InvoicePdfText.Buyer), _invoice.Buyer, theme)));

        var identity = new List<string[]> {
            new[] { Label(InvoicePdfText.Currency), _invoice.Currency, Label(InvoicePdfText.BuyerReference), _invoice.BuyerReference ?? "-" }
        };
        content.Table(identity, style: new PdfTableStyle {
            HeaderRowCount = 0,
            RowStripeFill = null,
            BorderColor = theme.Border,
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7D),
            CornerRadius = global::OfficeIMO.OfficeLength.Points(theme.CornerRadius),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(10D),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(7D),
            FontSize = global::OfficeIMO.OfficeLength.Points(9D),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(14D),
            BodyColumnFills = new List<global::OfficeIMO.OfficeColor?> { theme.Surface, null, theme.Surface, null },
            ColumnWidthWeights = new List<double> { 1.1D, 1.4D, 1.4D, 2.1D }
        });

        ComposeModernLines(content, theme);
        ComposeModernFinancialDetails(content, theme);
        ComposeModernTotals(content, theme);
        ComposeDetails(content, theme);
    }

    private void PartyCard(PdfContentBuilder content, string label, InvoiceParty party, InvoicePdfTheme theme) {
        content.Panel(card => {
            card.Paragraph(p => p.Bold(label.ToUpperInvariant(), theme.Accent), style: new PdfParagraphStyle {
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(4D)
            });
            card.Paragraph(p => p.Text(Party(party)), defaultColor: theme.Text, style: new PdfParagraphStyle {
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D),
                LineHeight = 1.25D
            });
        }, CardStyle(theme));
    }

    private void ComposeModernLines(PdfContentBuilder content, InvoicePdfTheme theme) {
        if (_invoice.Lines.Count == 0) return;
        content.Table(LineRows(), style: new PdfTableStyle {
            HeaderRowCount = 1,
            HeaderFill = theme.Text,
            HeaderTextColor = global::OfficeIMO.OfficeColor.White,
            HeaderFontSize = global::OfficeIMO.OfficeLength.Points(9D),
            FontSize = global::OfficeIMO.OfficeLength.Points(9D),
            BorderColor = theme.Border,
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5D),
            CornerRadius = global::OfficeIMO.OfficeLength.Points(theme.CornerRadius),
            RowStripeFill = theme.Surface,
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(7D),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(7D),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(12D),
            ColumnWidthWeights = LineColumnWeights(),
            Alignments = LineColumnAlignments()
        });
    }

    private void ComposeModernTotals(PdfContentBuilder content, InvoicePdfTheme theme) {
        var summary = new List<string[]>();
        void Add(InvoicePdfText label, decimal? amount) {
            if (amount.HasValue) summary.Add(new[] { Label(label), Money(amount.Value) });
        }
        void AddNonZero(InvoicePdfText label, decimal? amount) {
            if (amount.HasValue && amount.Value != 0M) summary.Add(new[] { Label(label), Money(amount.Value) });
        }
        // The summary keeps to the rows that change the amount, so it fits beside the line items on the first
        // page: the line total repeats the net total without allowances or charges, and nothing was prepaid.
        if (_amounts.Totals.LineNetTotal != _amounts.Totals.TaxExclusiveTotal) {
            Add(InvoicePdfText.LineNetTotal, _amounts.Totals.LineNetTotal);
        }
        AddNonZero(InvoicePdfText.Allowances, _amounts.Totals.AllowanceTotal);
        AddNonZero(InvoicePdfText.Charges, _amounts.Totals.ChargeTotal);
        Add(InvoicePdfText.TotalExcludingVat, _amounts.Totals.TaxExclusiveTotal);
        Add(InvoicePdfText.VatTotal, _amounts.Totals.TaxTotal);
        Add(InvoicePdfText.TotalIncludingVat, _amounts.Totals.TaxInclusiveTotal);
        AddNonZero(InvoicePdfText.Prepaid, _amounts.PrepaidAmount);
        AddNonZero(InvoicePdfText.Rounding, _amounts.RoundingAmount);
        summary.Add(new[] { Label(InvoicePdfText.AmountDue), Money(_amounts.PayableAmount) });

        string paymentReference = _invoice.Payments
            .Select(payment => payment.Reference)
            .FirstOrDefault(reference => reference != null) ?? _invoice.PaymentReference ?? _invoice.Number;
        int paymentSummaryLength = paymentReference.Length +
            (_invoice.PaymentTerms?.Length ?? 0);
        bool paymentSummaryIsMultiline = ContainsLineBreak(paymentReference) ||
            ContainsLineBreak(_invoice.PaymentTerms);
        content.Row(row => row
            .Style(new PdfRowStyle {
                Gap = global::OfficeIMO.OfficeLength.Points(18D),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(12D),
                KeepTogether = paymentSummaryLength <= 600 && !paymentSummaryIsMultiline
            })
            .PercentColumn(55D, column => {
                column.PanelParagraph(p => p
                    .Bold(Label(InvoicePdfText.PaymentReference) + "\n", theme.Text)
                    .Color(theme.MutedText)
                    .Text(paymentReference)
                    .Text(_invoice.PaymentTerms == null ? string.Empty : "\n" + _invoice.PaymentTerms),
                    CardStyle(theme));
            })
            .PercentColumn(45D, column => {
                column.Table(summary, style: new PdfTableStyle {
                    HeaderRowCount = 0,
                    FontSize = global::OfficeIMO.OfficeLength.Points(10D),
                    BorderColor = theme.Border,
                    BorderWidth = global::OfficeIMO.OfficeLength.Points(0.6D),
                    CornerRadius = global::OfficeIMO.OfficeLength.Points(theme.CornerRadius),
                    FooterRowCount = 1,
                    FooterFill = theme.Accent,
                    FooterTextColor = global::OfficeIMO.OfficeColor.White,
                    RowStripeFill = theme.Surface,
                    CellPaddingX = OfficeLength.Points(9D),
                    CellPaddingY = OfficeLength.Points(5D),
                    CellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign> {
                        [(summary.Count - 1, 1)] = PdfColumnAlign.Right
                    },
                    AutoFitColumns = true,
                    Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right }
                });
            }));
    }

    private static bool ContainsLineBreak(string? value) =>
        value?.IndexOf('\n') >= 0 || value?.IndexOf('\r') >= 0;

    private void ComposeModernFinancialDetails(PdfContentBuilder content, InvoicePdfTheme theme) {
        if (_invoice.AllowancesAndCharges.Count != 0) {
            var adjustments = _invoice.AllowancesAndCharges.Select(item => new[] {
                Label(item.IsCharge ? InvoicePdfText.Charge : InvoicePdfText.Allowance),
                AdjustmentDetails(item),
                VatCategory(item.Tax!.Code, item.Tax.Rate),
                Money(item.Amount)
            });
            content.H2(Label(InvoicePdfText.DocumentAdjustments), PdfAlign.Left, theme.Text, style: SectionHeadingStyle());
            content.Table(adjustments, style: ModernFinancialTableStyle(theme, false,
                PdfColumnAlign.Left, PdfColumnAlign.Left, PdfColumnAlign.Left, PdfColumnAlign.Right));
        }

        if (_amounts.Taxes.Count != 0) {
            var taxes = new List<string[]> {
                new[] { Label(InvoicePdfText.VatCategory), Label(InvoicePdfText.TaxableAmount), Label(InvoicePdfText.VatAmount), Label(InvoicePdfText.Exemption) }
            };
            taxes.AddRange(_amounts.Taxes.Select(tax => new[] {
                VatCategory(tax.CategoryCode, tax.Rate),
                Money(tax.TaxableAmount),
                Money(tax.TaxAmount),
                Join(tax.ExemptionReasonCode, tax.ExemptionReason)
            }));
            content.Table(taxes, style: ModernFinancialTableStyle(theme, true,
                PdfColumnAlign.Left, PdfColumnAlign.Right, PdfColumnAlign.Right, PdfColumnAlign.Left));
        }
    }

    private static PdfTableStyle ModernFinancialTableStyle(InvoicePdfTheme theme, bool hasHeader, params PdfColumnAlign[] alignments) => new PdfTableStyle {
        HeaderRowCount = hasHeader ? 1 : 0,
        HeaderFill = theme.Text,
        HeaderTextColor = OfficeColor.White,
        HeaderFontSize = OfficeLength.Points(9D),
        FontSize = OfficeLength.Points(9D),
        BorderColor = theme.Border,
        BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5D),
        CornerRadius = global::OfficeIMO.OfficeLength.Points(theme.CornerRadius),
        RowStripeFill = theme.Surface,
        CellPaddingX = OfficeLength.Points(7D),
        CellPaddingY = OfficeLength.Points(6D),
        SpacingAfter = OfficeLength.Points(12D),
        ColumnWidthWeights = new List<double> { 1.6D, 1.4D, 1.4D, 1.6D },
        AutoFitColumns = true,
        Alignments = new List<PdfColumnAlign>(alignments)
    };

    private void ComposeApprovals(PdfContentBuilder content, InvoicePdfTheme theme) {
        if (_layout.Approvals.Count == 0) return;
        content.H2(Label(InvoicePdfText.Approvals), PdfAlign.Left, theme.Text, style: SectionHeadingStyle());
        content.Row(row => {
            row.Style(new PdfRowStyle { Gap = global::OfficeIMO.OfficeLength.Points(12D) });
            foreach (InvoicePdfApproval approval in _layout.Approvals) {
                row.RelativeColumn(column => column.Panel(card => {
                    card.Paragraph(p => p.Text(approval.Label.ToUpperInvariant()), defaultColor: theme.MutedText, style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(9D) });
                    card.HR(global::OfficeIMO.OfficeLength.Points(0.8D), theme.Border, global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(6D));
                    card.Paragraph(p => p.Bold(approval.Name, theme.Text));
                    string detail = string.Join(" · ", new[] {
                        approval.Role,
                        approval.Date.HasValue ? Date(approval.Date) : null
                    }.Where(value => !string.IsNullOrWhiteSpace(value)));
                    if (detail.Length > 0) card.Paragraph(p => p.Text(detail), defaultColor: theme.MutedText, style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) });
                }, CardStyle(theme)), 1D);
            }
        });
    }

    private static PdfPanelStyle CardStyle(InvoicePdfTheme theme) => new PdfPanelStyle {
        Background = theme.Surface,
        BorderColor = theme.Border,
        BorderWidth = global::OfficeIMO.OfficeLength.Points(0.6D),
        CornerRadius = global::OfficeIMO.OfficeLength.Points(theme.CornerRadius),
        PaddingX = global::OfficeIMO.OfficeLength.Points(10D),
        PaddingY = global::OfficeIMO.OfficeLength.Points(9D),
        SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D)
    };
}
