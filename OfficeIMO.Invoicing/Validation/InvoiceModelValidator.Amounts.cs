namespace OfficeIMO.Invoicing;

public static partial class InvoiceModelValidator {
    private sealed partial class ModelChecks {
        internal bool AggregateAmounts(Invoice invoice, InvoiceProfile profile) {
            InvoiceDeclaredTotals? totals = invoice.DeclaredTotals;
            if (totals == null) {
                Error("INV-REQUIRED", "Aggregate-only Factur-X data requires declared totals.", "DeclaredTotals");
                return false;
            }
            bool valid = true;
            decimal? Required(decimal? value, string path) {
                if (!value.HasValue) {
                    Error("INV-REQUIRED", "The retained Factur-X profile requires this declared amount.", path);
                    valid = false;
                } else Money(value.Value, path);
                return value;
            }
            decimal? line = profile == InvoiceProfile.BasicWithoutLines
                ? Required(totals.LineNetTotal, "DeclaredTotals.LineNetTotal")
                : totals.LineNetTotal;
            decimal? taxExclusive = Required(totals.TaxExclusiveTotal, "DeclaredTotals.TaxExclusiveTotal");
            decimal? tax = Required(totals.TaxTotal, "DeclaredTotals.TaxTotal");
            decimal? taxInclusive = Required(totals.TaxInclusiveTotal, "DeclaredTotals.TaxInclusiveTotal");
            decimal? payable = Required(totals.PayableAmount, "DeclaredTotals.PayableAmount");
            if (totals.AllowanceTotal.HasValue) Money(totals.AllowanceTotal.Value, "DeclaredTotals.AllowanceTotal");
            if (totals.ChargeTotal.HasValue) Money(totals.ChargeTotal.Value, "DeclaredTotals.ChargeTotal");
            Money(invoice.PrepaidAmount, "PrepaidAmount");
            Money(invoice.RoundingAmount, "RoundingAmount");
            if (profile == InvoiceProfile.BasicWithoutLines && invoice.DeclaredTaxes.Count == 0) {
                Error("INV-REQUIRED", "Factur-X BASIC WL requires a declared VAT breakdown.", "DeclaredTaxes");
                valid = false;
            }
            if (!valid) return false;
            if (line.HasValue)
                Compare(taxExclusive, InvoiceArithmetic.Sum(new[] { line.Value, -(totals.AllowanceTotal ?? 0m), totals.ChargeTotal ?? 0m }),
                    "DeclaredTotals.TaxExclusiveTotal");
            Compare(taxInclusive, InvoiceArithmetic.Add(taxExclusive!.Value, tax!.Value), "DeclaredTotals.TaxInclusiveTotal");
            Compare(payable, InvoiceArithmetic.Sum(new[] { taxInclusive!.Value, -invoice.PrepaidAmount, invoice.RoundingAmount }),
                "DeclaredTotals.PayableAmount");
            if (invoice.DeclaredTaxes.Count != 0) {
                var declaredKeys = new HashSet<(string Code, decimal? Rate)>();
                for (int index = 0; index < invoice.DeclaredTaxes.Count; index++) {
                    InvoiceDeclaredTax declared = invoice.DeclaredTaxes[index];
                    string categoryPath = "DeclaredTaxes[" + index + "].Category";
                    Tax(declared.Category, categoryPath, true);
                    if (!declaredKeys.Add((declared.Category.Code, InvoiceCalculator.NormalizeRate(declared.Category)))) {
                        Error("INV-TAX-BREAKDOWN", "VAT category/rate is declared more than once.", categoryPath);
                        valid = false;
                    }
                    Money(declared.TaxableAmount, "DeclaredTaxes.TaxableAmount");
                    Money(declared.TaxAmount, "DeclaredTaxes.TaxAmount");
                    Compare(declared.TaxAmount, InvoiceArithmetic.RoundedProduct(
                        declared.TaxableAmount, InvoiceCalculator.NormalizeRate(declared.Category) ?? 0m, 100m),
                        "DeclaredTaxes[" + index + "].TaxAmount", 0.01m);
                }
                Compare(taxExclusive, InvoiceArithmetic.Sum(invoice.DeclaredTaxes.Select(item => item.TaxableAmount)),
                    "DeclaredTotals.TaxExclusiveTotal");
                Compare(tax, InvoiceArithmetic.Sum(invoice.DeclaredTaxes.Select(item => item.TaxAmount)),
                    "DeclaredTotals.TaxTotal");
            }
            return valid;
        }

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
