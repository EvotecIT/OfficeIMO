namespace OfficeIMO.Invoicing.Pdf;

internal sealed class PdfInvoiceAmounts {
    private PdfInvoiceAmounts(
        IReadOnlyList<decimal> lines,
        IReadOnlyList<PdfInvoiceTaxAmount> taxes,
        InvoiceDeclaredTotals totals,
        decimal prepaid,
        decimal rounding) {
        Lines = lines;
        Taxes = taxes;
        Totals = totals;
        PrepaidAmount = prepaid;
        RoundingAmount = rounding;
    }

    internal IReadOnlyList<decimal> Lines { get; }
    internal IReadOnlyList<PdfInvoiceTaxAmount> Taxes { get; }
    internal InvoiceDeclaredTotals Totals { get; }
    internal decimal PrepaidAmount { get; }
    internal decimal RoundingAmount { get; }
    internal decimal PayableAmount => Totals.PayableAmount ?? throw new InvalidDataException("The retained invoice profile has no payable amount.");

    internal static PdfInvoiceAmounts Create(Invoice invoice, InvoiceProfile profile) {
        if (profile is not InvoiceProfile.Minimum and not InvoiceProfile.BasicWithoutLines) {
            InvoiceModelValidationResult validation = InvoiceModelValidator.Validate(invoice);
            validation.ThrowIfInvalid();
            InvoiceCalculation calculation = validation.Calculation!;
            return new PdfInvoiceAmounts(
                calculation.Lines.Select(line => line.NetAmount).ToArray(),
                calculation.Taxes.Select(tax => new PdfInvoiceTaxAmount(tax.CategoryCode, tax.Rate, tax.TaxableAmount,
                    tax.TaxAmount, tax.ExemptionReason, tax.ExemptionReasonCode)).ToArray(),
                new InvoiceDeclaredTotals {
                    LineNetTotal = calculation.LineNetTotal,
                    AllowanceTotal = calculation.AllowanceTotal,
                    ChargeTotal = calculation.ChargeTotal,
                    TaxExclusiveTotal = calculation.TaxExclusiveTotal,
                    TaxTotal = calculation.TaxTotal,
                    TaxInclusiveTotal = calculation.TaxInclusiveTotal,
                    PayableAmount = calculation.PayableAmount
                },
                calculation.PrepaidAmount,
                calculation.RoundingAmount);
        }

        InvoiceDeclaredTotals totals = invoice.DeclaredTotals
            ?? throw new InvalidDataException("The retained invoice profile has no monetary totals.");
        if (!totals.TaxExclusiveTotal.HasValue || !totals.TaxTotal.HasValue || !totals.TaxInclusiveTotal.HasValue || !totals.PayableAmount.HasValue)
            throw new InvalidDataException("The retained invoice profile is missing a required monetary total.");
        return new PdfInvoiceAmounts(
            Array.Empty<decimal>(),
            invoice.DeclaredTaxes.Select(tax => new PdfInvoiceTaxAmount(tax.Category.Code, tax.Category.Rate,
                tax.TaxableAmount, tax.TaxAmount, tax.Category.ExemptionReason, tax.Category.ExemptionReasonCode)).ToArray(),
            totals,
            invoice.PrepaidAmount,
            invoice.RoundingAmount);
    }
}

internal sealed class PdfInvoiceTaxAmount {
    internal PdfInvoiceTaxAmount(string categoryCode, decimal? rate, decimal taxableAmount, decimal taxAmount, string? exemptionReason, string? exemptionReasonCode) {
        CategoryCode = categoryCode;
        Rate = rate;
        TaxableAmount = taxableAmount;
        TaxAmount = taxAmount;
        ExemptionReason = exemptionReason;
        ExemptionReasonCode = exemptionReasonCode;
    }

    internal string CategoryCode { get; }
    internal decimal? Rate { get; }
    internal decimal TaxableAmount { get; }
    internal decimal TaxAmount { get; }
    internal string? ExemptionReason { get; }
    internal string? ExemptionReasonCode { get; }
}
