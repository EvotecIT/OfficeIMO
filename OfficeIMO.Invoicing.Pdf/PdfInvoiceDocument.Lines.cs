using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class PdfInvoiceDocument {
    private List<string[]> LineRows() {
        var rows = new List<string[]> { _layout.LineColumns.Select(column => Label(ColumnLabel(column))).ToArray() };
        for (int index = 0; index < _invoice.Lines.Count; index++) {
            InvoiceLine line = _invoice.Lines[index];
            int rowIndex = index;
            rows.Add(_layout.LineColumns.Select(column => LineCell(column, line, rowIndex)).ToArray());
        }
        return rows;
    }

    private string LineCell(InvoicePdfLineColumn column, InvoiceLine line, int index) => column switch {
        InvoicePdfLineColumn.Item => LineText(line, selectedColumns: true),
        InvoicePdfLineColumn.Quantity => NumberText(line.Quantity) + (_layout.LineColumns.Contains(InvoicePdfLineColumn.Unit) ? string.Empty : " " + Unit(line.UnitCode)),
        InvoicePdfLineColumn.NetPrice => Price(line.UnitPrice, line),
        InvoicePdfLineColumn.Vat => line.Tax.Code + (line.Tax.Rate.HasValue ? " " + NumberText(line.Tax.Rate.Value) + "%" : string.Empty),
        InvoicePdfLineColumn.NetAmount => Money(_amounts.Lines[index]),
        InvoicePdfLineColumn.Unit => Unit(line.UnitCode),
        InvoicePdfLineColumn.LineIdentifier => line.Id,
        InvoicePdfLineColumn.Description => line.Description ?? string.Empty,
        InvoicePdfLineColumn.Period => Period(line.Period) ?? string.Empty,
        InvoicePdfLineColumn.SellerItem => line.SellerItemIdentifier ?? string.Empty,
        InvoicePdfLineColumn.BuyerItem => line.BuyerItemIdentifier ?? string.Empty,
        InvoicePdfLineColumn.StandardItem => Identifier(string.Empty, line.StandardItemIdentifier)?.TrimStart(':', ' ') ?? string.Empty,
        InvoicePdfLineColumn.AccountingReference => line.AccountingReference ?? string.Empty,
        InvoicePdfLineColumn.GrossPrice => line.GrossPrice.HasValue ? Price(line.GrossPrice.Value, line) : string.Empty,
        InvoicePdfLineColumn.PriceDiscount => line.PriceDiscount.HasValue ? Price(line.PriceDiscount.Value, line) : string.Empty,
        _ => throw new InvalidOperationException("Invoice column is undefined.")
    };

    private string Price(decimal value, InvoiceLine line) => NumberText(value) + " / " + NumberText(line.PriceBaseQuantity) + " " + Unit(line.UnitCode);
    private static InvoicePdfText ColumnLabel(InvoicePdfLineColumn column) => column switch {
        InvoicePdfLineColumn.Item => InvoicePdfText.Item,
        InvoicePdfLineColumn.Quantity => InvoicePdfText.Quantity,
        InvoicePdfLineColumn.NetPrice => InvoicePdfText.NetPrice,
        InvoicePdfLineColumn.Vat => InvoicePdfText.Vat,
        InvoicePdfLineColumn.NetAmount => InvoicePdfText.NetAmount,
        InvoicePdfLineColumn.Unit => InvoicePdfText.Unit,
        InvoicePdfLineColumn.LineIdentifier => InvoicePdfText.LineIdentifier,
        InvoicePdfLineColumn.Description => InvoicePdfText.Description,
        InvoicePdfLineColumn.Period => InvoicePdfText.Period,
        InvoicePdfLineColumn.SellerItem => InvoicePdfText.SellerItem,
        InvoicePdfLineColumn.BuyerItem => InvoicePdfText.BuyerItem,
        InvoicePdfLineColumn.StandardItem => InvoicePdfText.StandardItem,
        InvoicePdfLineColumn.AccountingReference => InvoicePdfText.AccountingReference,
        InvoicePdfLineColumn.GrossPrice => InvoicePdfText.GrossPrice,
        InvoicePdfLineColumn.PriceDiscount => InvoicePdfText.PriceDiscount,
        _ => throw new InvalidOperationException("Invoice column is undefined.")
    };

    private List<double> LineColumnWeights() => _layout.LineColumns.Select(column => column switch {
        InvoicePdfLineColumn.Item => 3.7D,
        InvoicePdfLineColumn.Description => 2.5D,
        InvoicePdfLineColumn.Period => 2D,
        InvoicePdfLineColumn.Unit => 1.7D,
        InvoicePdfLineColumn.NetPrice or InvoicePdfLineColumn.NetAmount or InvoicePdfLineColumn.GrossPrice or InvoicePdfLineColumn.PriceDiscount => 1.7D,
        _ => 1.3D
    }).ToList();

    private List<PdfColumnAlign> LineColumnAlignments() => _layout.LineColumns.Select(column => column is
        InvoicePdfLineColumn.Quantity or InvoicePdfLineColumn.NetPrice or InvoicePdfLineColumn.NetAmount or
        InvoicePdfLineColumn.Vat or InvoicePdfLineColumn.GrossPrice or InvoicePdfLineColumn.PriceDiscount
        ? PdfColumnAlign.Right : PdfColumnAlign.Left).ToList();

    private string Unit(string code) => CodeDescription(code, _layout.UnitCodeDisplay, language => language.GetUnitDescription(code));
    private string PaymentMeans(InvoicePayment payment) => Join(CodeDescription(payment.MeansCode, _layout.PaymentCodeDisplay,
        language => language.GetPaymentDescription(payment.MeansCode)), payment.MeansText).Replace("\n", " ");
    private string CodeDescription(string code, InvoicePdfCodeDisplay mode, Func<InvoicePdfLanguagePack, string?> translate) {
        if (mode == InvoicePdfCodeDisplay.Code) return code;
        string description = string.Join(_layout.LabelSeparator, _layout.Languages.Select(translate)
            .Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal));
        if (description.Length == 0) return code;
        return mode == InvoicePdfCodeDisplay.Description ? description : code + " (" + description + ")";
    }
}
