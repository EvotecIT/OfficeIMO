using System.Globalization;
using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class PdfInvoiceDocument {
    private static readonly string[] PartyHeaders = { "Seller", "Buyer" };
    private static readonly string[] LineHeaders = { "Item", "Quantity", "Net price", "VAT", "Net amount" };
    private static readonly string[] TaxHeaders = { "VAT category", "Taxable amount", "VAT amount", "Exemption" };
    private static readonly CultureInfo FormatCulture = CultureInfo.InvariantCulture;
    private void Compose(PdfContentBuilder content) {
        content.H1(DocumentTitle);
        content.Table(new[] { PartyHeaders, new[] { Party(_invoice.Seller), Party(_invoice.Buyer) } },
            style: new PdfTableStyle { HeaderRowCount = 1, RowStripeFill = null, SpacingAfter = 8, FontSize = 9 });
        var identity = new List<string[]> {
            new[] { "Issued", Date(_invoice.IssueDate), "Due", Date(_invoice.DueDate) }
        };
        if (_invoice.BuyerReference != null) identity.Add(new[] { "Buyer reference", _invoice.BuyerReference, "Currency", _invoice.Currency });
        content.Table(identity, style: new PdfTableStyle { HeaderRowCount = 0, RowStripeFill = null, SpacingAfter = 14, FontSize = 9,
            ColumnWidthWeights = new List<double> { 1.2, 2.8, 1, 2 } });
        var lines = new List<string[]> { LineHeaders };
        for (int index = 0; index < _invoice.Lines.Count; index++) {
            InvoiceLine line = _invoice.Lines[index];
            lines.Add(new[] {
                LineText(line), NumberText(line.Quantity) + " " + line.UnitCode,
                NumberText(line.UnitPrice) + " / " + NumberText(line.PriceBaseQuantity) + " " + line.UnitCode,
                line.Tax.Code + (line.Tax.Rate.HasValue ? " " + NumberText(line.Tax.Rate.Value) + "%" : string.Empty),
                Money(_amounts.Lines[index].NetAmount)
            });
        }
        content.Table(lines, style: new PdfTableStyle {
            HeaderRowCount = 1, SpacingAfter = 12, FontSize = 9,
            ColumnWidthWeights = new List<double> { 3.7, 1.3, 1.7, 1.1, 1.7 },
            Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right, PdfColumnAlign.Right, PdfColumnAlign.Right, PdfColumnAlign.Right }
        });
        if (_invoice.AllowancesAndCharges.Count != 0) {
            TableGroup(content, "Document adjustments", _invoice.AllowancesAndCharges.Select(item => new[] {
                item.IsCharge ? "Charge" : "Allowance", AdjustmentDetails(item),
                item.Tax!.Code + (item.Tax.Rate.HasValue ? " " + NumberText(item.Tax.Rate.Value) + "%" : string.Empty), Money(item.Amount)
            }));
        }
        var taxes = new List<string[]> { TaxHeaders };
        taxes.AddRange(_amounts.Taxes.Select(tax => new[] {
            tax.CategoryCode + (tax.Rate.HasValue ? " " + NumberText(tax.Rate.Value) + "%" : string.Empty),
            Money(tax.TaxableAmount), Money(tax.TaxAmount), Join(tax.ExemptionReasonCode, tax.ExemptionReason)
        }));
        content.Table(taxes, style: new PdfTableStyle { HeaderRowCount = 1, FontSize = 9, SpacingAfter = 12 });
        content.Flow(summary => summary.Table(new[] {
            new[] { "Line net total", Money(_amounts.LineNetTotal) },
            new[] { "Allowances", Money(_amounts.AllowanceTotal) },
            new[] { "Charges", Money(_amounts.ChargeTotal) },
            new[] { "Total excluding VAT", Money(_amounts.TaxExclusiveTotal) },
            new[] { "VAT total", Money(_amounts.TaxTotal) },
            new[] { "Total including VAT", Money(_amounts.TaxInclusiveTotal) },
            new[] { "Prepaid", Money(_amounts.PrepaidAmount) },
            new[] { "Rounding", Money(_amounts.RoundingAmount) },
            new[] { "Amount due", Money(_amounts.PayableAmount) }
        }, style: new PdfTableStyle {
            HeaderRowCount = 0, FontSize = 10, SpacingAfter = 14,
            Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right }
        }), new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.MoveToNextPage });
        ComposeDetails(content);
    }
    private static PdfTableStyle PlainTable() => new PdfTableStyle { HeaderRowCount = 0, FontSize = 9, SpacingAfter = 10, RowStripeFill = null };
    private string Money(decimal value) => value.ToString("0.00", FormatCulture) + " " + _invoice.Currency;
    private static string NumberText(decimal value) => value.ToString("0.############################", FormatCulture);
    private static string Date(DateTime? value) => value?.ToString("yyyy-MM-dd", FormatCulture) ?? string.Empty;
    private static string? Period(InvoicePeriod? period) => period == null ? null :
        period.Start.HasValue && period.End.HasValue ? Date(period.Start) + " to " + Date(period.End) :
        period.Start.HasValue ? "From " + Date(period.Start) : period.End.HasValue ? "Until " + Date(period.End) : null;
    private static string? Identifier(string label, InvoiceIdentifier? identifier) => identifier == null ? null :
        label + (identifier.SchemeId == null ? string.Empty : " (" + identifier.SchemeId + ")") + ": " + identifier.Value;
    private static string Join(params string?[] values) => string.Join("\n", values.Where(value => !string.IsNullOrWhiteSpace(value)));
    private static string Address(InvoiceAddress address) => Join(address.Line1, address.Line2, address.Line3,
        Join(address.PostCode, address.City).Replace("\n", " "), address.Subdivision, address.CountryCode);
    private static string Party(InvoiceParty party) => Join(party.Name, party.TradingName, Address(party.Address),
        party.VatIdentifier == null ? null : "VAT: " + party.VatIdentifier,
        party.TaxRegistration == null ? null : "Tax registration: " + party.TaxRegistration,
        string.Join("\n", party.Identifiers.Select(identifier => Identifier("Identifier", identifier))),
        Identifier("Legal registration", party.LegalRegistration), Identifier("Electronic address", party.ElectronicAddress),
        party.LegalInformation, party.Contact?.Name, party.Contact?.Email, party.Contact?.Telephone);
    private string LineText(InvoiceLine line) => Join(line.Id + ". " + line.Name, line.Description, line.Note,
        line.OrderLineReference == null ? null : "Order line: " + line.OrderLineReference,
        line.AccountingReference == null ? null : "Accounting reference: " + line.AccountingReference,
        Identifier("Object", line.ObjectIdentifier), Identifier("Standard item", line.StandardItemIdentifier),
        line.SellerItemIdentifier == null ? null : "Seller item: " + line.SellerItemIdentifier,
        line.BuyerItemIdentifier == null ? null : "Buyer item: " + line.BuyerItemIdentifier,
        line.OriginCountryCode == null ? null : "Origin: " + line.OriginCountryCode,
        string.Join("\n", line.Classifications.Select(item => "Classification (" + item.ListId +
            (item.ListVersion == null ? string.Empty : ", version " + item.ListVersion) + "): " + item.Value)),
        string.Join("\n", line.Attributes.Select(item => item.Name + ": " + item.Value)),
        Period(line.Period) is string period ? "Period: " + period : null,
        line.GrossPrice.HasValue ? "Gross price: " + NumberText(line.GrossPrice.Value) + " " + _invoice.Currency + " / " + NumberText(line.PriceBaseQuantity) + " " + line.UnitCode : null,
        line.PriceDiscount.HasValue ? "Price discount: " + NumberText(line.PriceDiscount.Value) + " " + _invoice.Currency + " / " + NumberText(line.PriceBaseQuantity) + " " + line.UnitCode : null,
        line.AllowancesAndCharges.Count == 0 ? null : string.Join("\n", line.AllowancesAndCharges.Select(item =>
            (item.IsCharge ? "Charge: " : "Allowance: ") + Money(item.Amount) + "\n" + AdjustmentDetails(item))));
    private string AdjustmentDetails(InvoiceAllowanceCharge item) => Join(
        item.Reason, item.ReasonCode == null ? null : "Reason code: " + item.ReasonCode,
        item.BaseAmount.HasValue ? "Base: " + Money(item.BaseAmount.Value) : null,
        item.Percentage.HasValue ? "Percentage: " + NumberText(item.Percentage.Value) + "%" : null);
}
