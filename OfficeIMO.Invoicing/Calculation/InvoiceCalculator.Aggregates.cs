namespace OfficeIMO.Invoicing;

public static partial class InvoiceCalculator {
    private static InvoiceCalculation RecalculateAggregate(Invoice invoice) {
        InvoiceCalculation retained = FromDeclaredAggregate(invoice);
        // No line formula can be recovered from a reduced profile. Retain the declared
        // bases and VAT (including permitted source rounding), and rebuild dependent totals.
        decimal exclusive = retained.Taxes.Count == 0 ? retained.TaxExclusiveTotal
            : InvoiceArithmetic.Sum(retained.Taxes.Select(tax => tax.TaxableAmount));
        decimal vat = retained.Taxes.Count == 0 ? retained.TaxTotal
            : InvoiceArithmetic.Sum(retained.Taxes.Select(tax => tax.TaxAmount));
        decimal inclusive = InvoiceArithmetic.Add(exclusive, vat);
        decimal lineTotal = invoice.DeclaredTotals!.LineNetTotal.HasValue
            ? InvoiceArithmetic.Sum(new[] { exclusive, retained.AllowanceTotal, -retained.ChargeTotal })
            : exclusive;
        return new InvoiceCalculation(retained.Lines, retained.Taxes, lineTotal, retained.AllowanceTotal,
            retained.ChargeTotal, exclusive, vat, inclusive, invoice.PrepaidAmount, invoice.RoundingAmount,
            InvoiceArithmetic.Sum(new[] { inclusive, -invoice.PrepaidAmount, invoice.RoundingAmount }));
    }
}
