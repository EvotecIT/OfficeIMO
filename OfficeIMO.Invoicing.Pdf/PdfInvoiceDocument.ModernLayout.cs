using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class PdfInvoiceDocument {
    private void ComposeModern(PdfContentBuilder content, InvoicePdfTheme theme) {
        content.Row(row => {
            row.Style(new PdfRowStyle { Gap = 20D, SpacingAfter = 16D })
                .PercentColumn(44D, column => {
                    if (_layout.LogoBytesSnapshot != null) {
                        column.Image(
                            _layout.LogoBytesSnapshot,
                            _layout.LogoWidth,
                            _layout.LogoHeight,
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
                        column.H2(_invoice.Seller.Name, PdfAlign.Left, theme.Text);
                    }
                })
                .PercentColumn(56D, column => {
                    column.Paragraph(p => p.FontSize(18D).Bold(DocumentTitle, theme.Text), PdfAlign.Right, style: new PdfParagraphStyle {
                        SpacingBefore = 0D,
                        SpacingAfter = 6D,
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
            .Style(new PdfRowStyle { Gap = 12D, SpacingAfter = 12D })
            .PercentColumn(50D, column => PartyCard(column, Label(InvoicePdfText.Seller), _invoice.Seller, theme))
            .PercentColumn(50D, column => PartyCard(column, Label(InvoicePdfText.Buyer), _invoice.Buyer, theme)));

        var identity = new List<string[]> {
            new[] { Label(InvoicePdfText.Currency), _invoice.Currency, Label(InvoicePdfText.BuyerReference), _invoice.BuyerReference ?? "-" }
        };
        content.Table(identity, style: new PdfTableStyle {
            HeaderRowCount = 0,
            RowStripeFill = null,
            BorderColor = theme.Border,
            BorderWidth = 0.7D,
            CornerRadius = theme.CornerRadius,
            CellPaddingX = 10D,
            CellPaddingY = 7D,
            FontSize = 9D,
            SpacingAfter = 14D,
            BodyColumnFills = new List<PdfColor?> { theme.Surface, null, theme.Surface, null },
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
                SpacingBefore = 0D,
                SpacingAfter = 4D
            });
            card.Paragraph(p => p.Text(Party(party)), defaultColor: theme.Text, style: new PdfParagraphStyle {
                SpacingBefore = 0D,
                SpacingAfter = 0D,
                LineHeight = 1.25D
            });
        }, CardStyle(theme));
    }

    private void ComposeModernLines(PdfContentBuilder content, InvoicePdfTheme theme) {
        if (_invoice.Lines.Count == 0) return;
        var rows = new List<string[]> {
            new[] { Label(InvoicePdfText.Item), Label(InvoicePdfText.Quantity), Label(InvoicePdfText.NetPrice), Label(InvoicePdfText.Vat), Label(InvoicePdfText.NetAmount) }
        };
        for (int index = 0; index < _invoice.Lines.Count; index++) {
            InvoiceLine line = _invoice.Lines[index];
            rows.Add(new[] {
                LineText(line),
                NumberText(line.Quantity) + " " + line.UnitCode,
                NumberText(line.UnitPrice) + " / " + NumberText(line.PriceBaseQuantity) + " " + line.UnitCode,
                line.Tax.Code + (line.Tax.Rate.HasValue ? " " + NumberText(line.Tax.Rate.Value) + "%" : string.Empty),
                Money(_amounts.Lines[index])
            });
        }
        content.Table(rows, style: new PdfTableStyle {
            HeaderRowCount = 1,
            HeaderFill = theme.Text,
            HeaderTextColor = PdfColor.White,
            HeaderFontSize = 9D,
            FontSize = 9D,
            BorderColor = theme.Border,
            BorderWidth = 0.5D,
            CornerRadius = theme.CornerRadius,
            RowStripeFill = theme.Surface,
            CellPaddingX = 7D,
            CellPaddingY = 7D,
            SpacingAfter = 12D,
            ColumnWidthWeights = new List<double> { 3.7D, 1.15D, 1.55D, 1D, 1.6D },
            Alignments = new List<PdfColumnAlign> {
                PdfColumnAlign.Left,
                PdfColumnAlign.Right,
                PdfColumnAlign.Right,
                PdfColumnAlign.Right,
                PdfColumnAlign.Right
            }
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
        Add(InvoicePdfText.LineNetTotal, _amounts.Totals.LineNetTotal);
        AddNonZero(InvoicePdfText.Allowances, _amounts.Totals.AllowanceTotal);
        AddNonZero(InvoicePdfText.Charges, _amounts.Totals.ChargeTotal);
        Add(InvoicePdfText.TotalExcludingVat, _amounts.Totals.TaxExclusiveTotal);
        Add(InvoicePdfText.VatTotal, _amounts.Totals.TaxTotal);
        Add(InvoicePdfText.TotalIncludingVat, _amounts.Totals.TaxInclusiveTotal);
        Add(InvoicePdfText.Prepaid, _amounts.PrepaidAmount);
        AddNonZero(InvoicePdfText.Rounding, _amounts.RoundingAmount);
        summary.Add(new[] { Label(InvoicePdfText.AmountDue), Money(_amounts.PayableAmount) });

        string paymentReference = _invoice.Payments
            .Select(payment => payment.Reference)
            .FirstOrDefault(reference => reference != null) ?? _invoice.Number;
        int paymentSummaryLength = paymentReference.Length +
            (_invoice.PaymentTerms?.Length ?? 0);
        bool paymentSummaryIsMultiline = ContainsLineBreak(paymentReference) ||
            ContainsLineBreak(_invoice.PaymentTerms);
        content.Row(row => row
            .Style(new PdfRowStyle {
                Gap = 18D,
                SpacingAfter = 12D,
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
                    FontSize = 10D,
                    BorderColor = theme.Border,
                    BorderWidth = 0.6D,
                    CornerRadius = theme.CornerRadius,
                    FooterRowCount = 1,
                    FooterFill = theme.Accent,
                    FooterTextColor = PdfColor.White,
                    RowStripeFill = theme.Surface,
                    CellPaddingX = 9D,
                    CellPaddingY = 7D,
                    CellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign> {
                        [(summary.Count - 1, 1)] = PdfColumnAlign.Right
                    },
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
                item.Tax!.Code + (item.Tax.Rate.HasValue ? " " + NumberText(item.Tax.Rate.Value) + "%" : string.Empty),
                Money(item.Amount)
            });
            content.H2(Label(InvoicePdfText.DocumentAdjustments), PdfAlign.Left, theme.Text);
            content.Table(adjustments, style: ModernFinancialTableStyle(theme, false));
        }

        if (_amounts.Taxes.Count != 0) {
            var taxes = new List<string[]> {
                new[] { Label(InvoicePdfText.VatCategory), Label(InvoicePdfText.TaxableAmount), Label(InvoicePdfText.VatAmount), Label(InvoicePdfText.Exemption) }
            };
            taxes.AddRange(_amounts.Taxes.Select(tax => new[] {
                tax.CategoryCode + (tax.Rate.HasValue ? " " + NumberText(tax.Rate.Value) + "%" : string.Empty),
                Money(tax.TaxableAmount),
                Money(tax.TaxAmount),
                Join(tax.ExemptionReasonCode, tax.ExemptionReason)
            }));
            content.Table(taxes, style: ModernFinancialTableStyle(theme, true));
        }
    }

    private static PdfTableStyle ModernFinancialTableStyle(InvoicePdfTheme theme, bool hasHeader) => new PdfTableStyle {
        HeaderRowCount = hasHeader ? 1 : 0,
        HeaderFill = theme.Text,
        HeaderTextColor = PdfColor.White,
        HeaderFontSize = 8.5D,
        FontSize = 8.5D,
        BorderColor = theme.Border,
        BorderWidth = 0.5D,
        CornerRadius = theme.CornerRadius,
        RowStripeFill = theme.Surface,
        CellPaddingX = 7D,
        CellPaddingY = 6D,
        SpacingAfter = 12D,
        ColumnWidthWeights = new List<double> { 1.2D, 2.8D, 1.2D, 1.5D },
        Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Left,
            PdfColumnAlign.Left,
            PdfColumnAlign.Right,
            PdfColumnAlign.Right
        }
    };

    private void ComposeApprovals(PdfContentBuilder content, InvoicePdfTheme theme) {
        if (_layout.Approvals.Count == 0) return;
        content.H2(Label(InvoicePdfText.Approvals), PdfAlign.Left, theme.Text);
        content.Row(row => {
            row.Style(new PdfRowStyle { Gap = 12D });
            foreach (InvoicePdfApproval approval in _layout.Approvals) {
                row.RelativeColumn(column => column.Panel(card => {
                    card.Paragraph(p => p.Text(approval.Label.ToUpperInvariant()), defaultColor: theme.MutedText, style: new PdfParagraphStyle { SpacingAfter = 9D });
                    card.HR(0.8D, theme.Border, 0D, 6D);
                    card.Paragraph(p => p.Bold(approval.Name, theme.Text));
                    string detail = string.Join(" · ", new[] {
                        approval.Role,
                        approval.Date.HasValue ? Date(approval.Date) : null
                    }.Where(value => !string.IsNullOrWhiteSpace(value)));
                    if (detail.Length > 0) card.Paragraph(p => p.Text(detail), defaultColor: theme.MutedText, style: new PdfParagraphStyle { SpacingAfter = 0D });
                }, CardStyle(theme)), 1D);
            }
        });
    }

    private static PdfPanelStyle CardStyle(InvoicePdfTheme theme) => new PdfPanelStyle {
        Background = theme.Surface,
        BorderColor = theme.Border,
        BorderWidth = 0.6D,
        CornerRadius = theme.CornerRadius,
        PaddingX = 10D,
        PaddingY = 9D,
        SpacingAfter = 0D
    };
}
