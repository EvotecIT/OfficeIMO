namespace OfficeIMO.Invoicing;

public static partial class InvoiceModelValidator {
    private sealed partial class ModelChecks {
        internal void DeclaredAmounts(Invoice invoice, InvoiceCalculation calculation) {
            for (int index = 0; index < invoice.Lines.Count; index++)
                Compare(invoice.Lines[index].DeclaredNetAmount, calculation.Lines[index].FormulaNetAmount, "Lines[" + index + "].DeclaredNetAmount", 0.02m);
            InvoiceDeclaredTotals? totals = invoice.DeclaredTotals;
            if (totals != null) {
                Compare(totals.LineNetTotal, calculation.LineNetTotal, "DeclaredTotals.LineNetTotal");
                Compare(totals.AllowanceTotal, calculation.AllowanceTotal, "DeclaredTotals.AllowanceTotal");
                Compare(totals.ChargeTotal, calculation.ChargeTotal, "DeclaredTotals.ChargeTotal");
                Compare(totals.TaxExclusiveTotal, calculation.TaxExclusiveTotal, "DeclaredTotals.TaxExclusiveTotal");
                Compare(totals.TaxTotal, calculation.TaxTotal, "DeclaredTotals.TaxTotal");
                Compare(totals.TaxInclusiveTotal, calculation.TaxInclusiveTotal, "DeclaredTotals.TaxInclusiveTotal");
                Compare(totals.PayableAmount, calculation.PayableAmount, "DeclaredTotals.PayableAmount");
            }
            var matched = new HashSet<InvoiceCalculatedTax>();
            var expectedByCategory = calculation.Taxes.ToDictionary(tax => (tax.CategoryCode, tax.Rate));
            foreach (InvoiceDeclaredTax declared in invoice.DeclaredTaxes) {
                Tax(declared.Category, "DeclaredTaxes.Category", true);
                if (!expectedByCategory.TryGetValue((declared.Category.Code, InvoiceCalculator.NormalizeRate(declared.Category)), out InvoiceCalculatedTax? expected)) {
                    Error("INV-TAX-BREAKDOWN", "Declared VAT category/rate has no matching taxable amounts.", "DeclaredTaxes"); continue;
                }
                if (!matched.Add(expected)) Error("INV-TAX-BREAKDOWN", "VAT category/rate is declared more than once.", "DeclaredTaxes");
                Compare(declared.TaxableAmount, expected.TaxableAmount, "DeclaredTaxes.TaxableAmount");
                Compare(declared.TaxAmount, expected.FormulaTaxAmount, "DeclaredTaxes.TaxAmount", 0.01m);
            }
            if (invoice.DeclaredTaxes.Count != 0 && matched.Count != calculation.Taxes.Count)
                Error("INV-TAX-BREAKDOWN", "Source VAT breakdown omits a calculated category/rate.", "DeclaredTaxes");
        }
        private void Compare(decimal? declared, decimal expected, string path, decimal tolerance = 0m) {
            if (!declared.HasValue) return;
            Money(declared.Value, path);
            if (Math.Abs(declared.Value - expected) > tolerance) Error("INV-DECLARED-AMOUNT", "Declared " + declared.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " differs from calculated " + expected.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", path);
        }
    }
}
