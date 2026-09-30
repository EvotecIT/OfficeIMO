namespace OfficeIMO.Invoicing;

/// <summary>Result of updating monetary declarations after an edit, including any required accounting-currency refresh.</summary>
public sealed class InvoiceEditResult {
    internal InvoiceEditResult(InvoiceCalculation? calculation, IReadOnlyList<InvoiceDiagnostic> diagnostics) {
        Calculation = calculation; Diagnostics = diagnostics;
    }
    /// <summary>Updated amounts, or null if recalculation failed before changing declarations.</summary>
    public InvoiceCalculation? Calculation { get; }
    /// <summary>Operation findings, including an explicit refresh requirement when foreign-currency VAT became stale.</summary>
    public IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; }
    /// <summary>True when recalculation completed. Check diagnostics and target validation before writing.</summary>
    public bool Succeeded => Calculation != null;
}

/// <summary>Reports the consequences of monetary edits without hiding a required accounting-currency refresh.</summary>
public static class InvoiceEditor {
    /// <summary>Updates invoice declarations and reports expected calculation failures. The caller supplies any tax-point exchange rate.</summary>
    public static InvoiceEditResult Recalculate(Invoice invoice, decimal? taxExchangeRate = null) {
        if (invoice == null) throw new ArgumentNullException(nameof(invoice));
        try {
            InvoiceCalculation calculation = InvoiceCalculator.UpdateDeclaredAmounts(invoice, taxExchangeRate);
            IReadOnlyList<InvoiceDiagnostic> diagnostics = invoice.TaxCurrency != null && !invoice.TaxAmountInAccountingCurrency.HasValue
                ? new[] { new InvoiceDiagnostic("INV-ACCOUNTING-VAT-REFRESH", "Invoice VAT changed or had no retained baseline. Supply a refreshed accounting-currency VAT amount or recalculate with an explicit tax-point exchange rate before writing.", "TaxAmountInAccountingCurrency", InvoiceDiagnosticSeverity.Warning) }
                : Array.Empty<InvoiceDiagnostic>();
            return new InvoiceEditResult(calculation, diagnostics);
        } catch (Exception exception) when (exception is ArgumentException || exception is OverflowException) {
            return new InvoiceEditResult(null, new[] { new InvoiceDiagnostic("INV-EDIT-CALCULATION", exception.Message, "Invoice") });
        }
    }
}
