using System.Globalization;
using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class PdfInvoiceDocument {
    private void Compose(PdfContentBuilder content) {
        content.H1(DocumentTitle);
        content.Table(new[] { new[] { Label(InvoicePdfText.Seller), Label(InvoicePdfText.Buyer) }, new[] { Party(_invoice.Seller), Party(_invoice.Buyer) } },
            style: new PdfTableStyle { HeaderRowCount = 1, RowStripeFill = null, SpacingAfter = 8, FontSize = 9 });
        var identity = new List<string[]> {
            new[] { Label(InvoicePdfText.Issued), Date(_invoice.IssueDate), Label(InvoicePdfText.Due), Date(_invoice.DueDate) }
        };
        if (_invoice.BuyerReference != null) identity.Add(new[] { Label(InvoicePdfText.BuyerReference), _invoice.BuyerReference, Label(InvoicePdfText.Currency), _invoice.Currency });
        content.Table(identity, style: new PdfTableStyle { HeaderRowCount = 0, RowStripeFill = null, SpacingAfter = 14, FontSize = 9,
            ColumnWidthWeights = new List<double> { 1.2, 2.8, 1, 2 } });
        var lines = new List<string[]> { new[] { Label(InvoicePdfText.Item), Label(InvoicePdfText.Quantity), Label(InvoicePdfText.NetPrice), Label(InvoicePdfText.Vat), Label(InvoicePdfText.NetAmount) } };
        for (int index = 0; index < _invoice.Lines.Count; index++) {
            InvoiceLine line = _invoice.Lines[index];
            lines.Add(new[] {
                LineText(line), NumberText(line.Quantity) + " " + line.UnitCode,
                NumberText(line.UnitPrice) + " / " + NumberText(line.PriceBaseQuantity) + " " + line.UnitCode,
                line.Tax.Code + (line.Tax.Rate.HasValue ? " " + NumberText(line.Tax.Rate.Value) + "%" : string.Empty),
                Money(_amounts.Lines[index])
            });
        }
        if (_invoice.Lines.Count != 0)
            content.Table(lines, style: new PdfTableStyle {
                HeaderRowCount = 1, SpacingAfter = 12, FontSize = 9,
                ColumnWidthWeights = new List<double> { 3.7, 1.3, 1.7, 1.1, 1.7 },
                Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right, PdfColumnAlign.Right, PdfColumnAlign.Right, PdfColumnAlign.Right }
            });
        if (_invoice.AllowancesAndCharges.Count != 0) {
            TableGroup(content, Label(InvoicePdfText.DocumentAdjustments), _invoice.AllowancesAndCharges.Select(item => new[] {
                Label(item.IsCharge ? InvoicePdfText.Charge : InvoicePdfText.Allowance), AdjustmentDetails(item),
                item.Tax!.Code + (item.Tax.Rate.HasValue ? " " + NumberText(item.Tax.Rate.Value) + "%" : string.Empty), Money(item.Amount)
            }));
        }
        var taxes = new List<string[]> { new[] { Label(InvoicePdfText.VatCategory), Label(InvoicePdfText.TaxableAmount), Label(InvoicePdfText.VatAmount), Label(InvoicePdfText.Exemption) } };
        taxes.AddRange(_amounts.Taxes.Select(tax => new[] {
            tax.CategoryCode + (tax.Rate.HasValue ? " " + NumberText(tax.Rate.Value) + "%" : string.Empty),
            Money(tax.TaxableAmount), Money(tax.TaxAmount), Join(tax.ExemptionReasonCode, tax.ExemptionReason)
        }));
        if (_amounts.Taxes.Count != 0)
            content.Table(taxes, style: new PdfTableStyle { HeaderRowCount = 1, FontSize = 9, SpacingAfter = 12 });
        var summaryRows = new List<string[]>();
        void AddSummary(InvoicePdfText label, decimal? value) {
            if (value.HasValue) summaryRows.Add(new[] { Label(label), Money(value.Value) });
        }
        AddSummary(InvoicePdfText.LineNetTotal, _amounts.Totals.LineNetTotal);
        AddSummary(InvoicePdfText.Allowances, _amounts.Totals.AllowanceTotal);
        AddSummary(InvoicePdfText.Charges, _amounts.Totals.ChargeTotal);
        AddSummary(InvoicePdfText.TotalExcludingVat, _amounts.Totals.TaxExclusiveTotal);
        AddSummary(InvoicePdfText.VatTotal, _amounts.Totals.TaxTotal);
        AddSummary(InvoicePdfText.TotalIncludingVat, _amounts.Totals.TaxInclusiveTotal);
        summaryRows.AddRange(new[] {
            new[] { Label(InvoicePdfText.Prepaid), Money(_amounts.PrepaidAmount) },
            new[] { Label(InvoicePdfText.Rounding), Money(_amounts.RoundingAmount) },
            new[] { Label(InvoicePdfText.AmountDue), Money(_amounts.PayableAmount) }
        });
        content.Flow(summary => summary.Table(summaryRows, style: new PdfTableStyle {
            HeaderRowCount = 0, FontSize = 10, SpacingAfter = 14,
            Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right }
        }), new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.MoveToNextPage });
        ComposeDetails(content);
    }
    private static PdfTableStyle PlainTable() => new PdfTableStyle { HeaderRowCount = 0, FontSize = 9, SpacingAfter = 10, RowStripeFill = null };
    private string Label(InvoicePdfText text) => string.Join(_layout.LabelSeparator,
        _layout.Languages.Select(language => language[text]).Distinct(StringComparer.Ordinal));
    private string Money(decimal value) => value.ToString("0.00", _layout.FormattingCulture) + " " + _invoice.Currency;
    private string NumberText(decimal value) => value.ToString("0.############################", _layout.FormattingCulture);
    private string Date(DateTime? value) => value?.ToString(_layout.DateFormat, _layout.FormattingCulture) ?? string.Empty;
    private string? Period(InvoicePeriod? period) => period == null ? null :
        period.Start.HasValue && period.End.HasValue ? Date(period.Start) + " " + Label(InvoicePdfText.To) + " " + Date(period.End) :
        period.Start.HasValue ? Label(InvoicePdfText.From) + " " + Date(period.Start) : period.End.HasValue ? Label(InvoicePdfText.Until) + " " + Date(period.End) : null;
    private static string? Identifier(string label, InvoiceIdentifier? identifier) => identifier == null ? null :
        label + (identifier.SchemeId == null ? string.Empty : " (" + identifier.SchemeId + ")") + ": " + identifier.Value;
    private static string Join(params string?[] values) => string.Join("\n", values.Where(value => !string.IsNullOrWhiteSpace(value)));
    private static string Address(InvoiceAddress address) => Join(address.Line1, address.Line2, address.Line3,
        Join(address.PostCode, address.City).Replace("\n", " "), address.Subdivision, address.CountryCode);
    private string Party(InvoiceParty party) => Join(party.Name, party.TradingName, Address(party.Address),
        string.Join("\n", party.TaxRegistrations.Select(registration => registration.SchemeId + ": " + registration.Identifier)),
        string.Join("\n", party.Identifiers.Select(identifier => Identifier(Label(InvoicePdfText.Identifier), identifier))),
        Identifier(Label(InvoicePdfText.LegalRegistration), party.LegalRegistration), Identifier(Label(InvoicePdfText.ElectronicAddress), party.ElectronicAddress),
        party.LegalInformation, party.Contact?.Name, party.Contact?.Email, party.Contact?.Telephone);
    private string LineText(InvoiceLine line) => Join(line.Id + ". " + line.Name, line.Description, line.Note,
        line.OrderLineReference == null ? null : Label(InvoicePdfText.OrderLine) + ": " + line.OrderLineReference,
        line.AccountingReference == null ? null : Label(InvoicePdfText.AccountingReference) + ": " + line.AccountingReference,
        Identifier(Label(InvoicePdfText.ObjectIdentifier), line.ObjectIdentifier), Identifier(Label(InvoicePdfText.StandardItem), line.StandardItemIdentifier),
        line.SellerItemIdentifier == null ? null : Label(InvoicePdfText.SellerItem) + ": " + line.SellerItemIdentifier,
        line.BuyerItemIdentifier == null ? null : Label(InvoicePdfText.BuyerItem) + ": " + line.BuyerItemIdentifier,
        line.OriginCountryCode == null ? null : Label(InvoicePdfText.Origin) + ": " + line.OriginCountryCode,
        string.Join("\n", line.Classifications.Select(item => Label(InvoicePdfText.Classification) + " (" + item.ListId +
            (item.ListVersion == null ? string.Empty : ", " + Label(InvoicePdfText.Version) + " " + item.ListVersion) + "): " + item.Value)),
        string.Join("\n", line.Attributes.Select(item => item.Name + ": " + item.Value)),
        Period(line.Period) is string period ? Label(InvoicePdfText.Period) + ": " + period : null,
        line.GrossPrice.HasValue ? Label(InvoicePdfText.GrossPrice) + ": " + NumberText(line.GrossPrice.Value) + " " + _invoice.Currency + " / " + NumberText(line.PriceBaseQuantity) + " " + line.UnitCode : null,
        line.PriceDiscount.HasValue ? Label(InvoicePdfText.PriceDiscount) + ": " + NumberText(line.PriceDiscount.Value) + " " + _invoice.Currency + " / " + NumberText(line.PriceBaseQuantity) + " " + line.UnitCode : null,
        line.AllowancesAndCharges.Count == 0 ? null : string.Join("\n", line.AllowancesAndCharges.Select(item =>
            Label(item.IsCharge ? InvoicePdfText.Charge : InvoicePdfText.Allowance) + ": " + Money(item.Amount) + "\n" + AdjustmentDetails(item))));
    private string AdjustmentDetails(InvoiceAllowanceCharge item) => Join(
        item.Reason, item.ReasonCode == null ? null : Label(InvoicePdfText.ReasonCode) + ": " + item.ReasonCode,
        item.BaseAmount.HasValue ? Label(InvoicePdfText.Base) + ": " + Money(item.BaseAmount.Value) : null,
        item.Percentage.HasValue ? Label(InvoicePdfText.Percentage) + ": " + NumberText(item.Percentage.Value) + "%" : null);
}
