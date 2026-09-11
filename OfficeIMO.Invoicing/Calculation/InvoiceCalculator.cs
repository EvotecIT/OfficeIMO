namespace OfficeIMO.Invoicing;

/// <summary>Shared decimal arithmetic for invoice XML, validation and visible presentation.</summary>
public static class InvoiceCalculator {
    /// <summary>Explicitly replaces source-declared amounts after editing prices, quantities, taxes or adjustments.</summary>
    public static InvoiceCalculation UpdateDeclaredAmounts(Invoice invoice) {
        InvoiceCalculation calculation = Calculate(invoice, false);
        for (int index = 0; index < invoice.Lines.Count; index++) invoice.Lines[index].DeclaredNetAmount = calculation.Lines[index].NetAmount;
        invoice.DeclaredTotals = new InvoiceDeclaredTotals {
            LineNetTotal = calculation.LineNetTotal, AllowanceTotal = calculation.AllowanceTotal, ChargeTotal = calculation.ChargeTotal,
            TaxExclusiveTotal = calculation.TaxExclusiveTotal, TaxTotal = calculation.TaxTotal, TaxInclusiveTotal = calculation.TaxInclusiveTotal, PayableAmount = calculation.PayableAmount
        };
        invoice.DeclaredTaxes.Clear();
        foreach (InvoiceCalculatedTax tax in calculation.Taxes) invoice.DeclaredTaxes.Add(new InvoiceDeclaredTax {
            Category = new InvoiceTaxCategory { Code = tax.CategoryCode, Rate = tax.Rate, ExemptionReason = tax.ExemptionReason, ExemptionReasonCode = tax.ExemptionReasonCode },
            TaxableAmount = tax.TaxableAmount, TaxAmount = tax.TaxAmount
        });
        return calculation;
    }
    /// <summary>Rounds to two decimal places using XPath round semantics (ties towards positive infinity).</summary>
    public static decimal RoundAmount(decimal value) {
        decimal rounded = decimal.Round(value, 2, MidpointRounding.AwayFromZero);
        // EN 16931 Schematron uses fn:round, including for negative invoice amounts.
        if (value < 0m && (value - decimal.Truncate(value)) % 0.01m == -0.005m) rounded += 0.01m;
        return rounded;
    }

    /// <summary>Calculates line totals, category-level VAT and the amount due. Throws on invalid arithmetic inputs or decimal overflow.</summary>
    public static InvoiceCalculation Calculate(Invoice invoice) => Calculate(invoice, true);

    private static InvoiceCalculation Calculate(Invoice invoice, bool preserveDeclaredAmounts) {
        if (invoice == null) throw new ArgumentNullException(nameof(invoice));
        var lines = new List<InvoiceCalculatedLine>();
        var groups = new Dictionary<(string Code, decimal? Rate), TaxGroup>();
        foreach (InvoiceLine line in invoice.Lines) {
            if (line == null) throw new ArgumentException("An invoice line is null.", nameof(invoice));
            if (line.PriceBaseQuantity <= 0m) throw new ArgumentException("Line price base quantity must be positive: " + line.Id, nameof(invoice));
            if (line.UnitPrice < 0m) throw new ArgumentException("Line net price cannot be negative: " + line.Id, nameof(invoice));
            decimal adjustments = 0m;
            foreach (InvoiceAllowanceCharge item in line.AllowancesAndCharges) {
                CheckAmount(item);
                adjustments = InvoiceArithmetic.Add(adjustments, item.IsCharge ? item.Amount : -item.Amount);
            }
            decimal net = InvoiceArithmetic.LineAmount(line.Quantity, line.UnitPrice, line.PriceBaseQuantity, adjustments, out decimal rounded);
            var calculated = new InvoiceCalculatedLine(line.Id, net, rounded, preserveDeclaredAmounts ? line.DeclaredNetAmount : null);
            lines.Add(calculated);
            AddTax(groups, line.Tax, calculated.NetAmount);
        }
        decimal allowances = 0m, charges = 0m;
        foreach (InvoiceAllowanceCharge item in invoice.AllowancesAndCharges) {
            CheckAmount(item);
            if (item.IsCharge) charges = InvoiceArithmetic.Add(charges, item.Amount); else allowances = InvoiceArithmetic.Add(allowances, item.Amount);
            AddTax(groups, item.Tax, item.IsCharge ? item.Amount : -item.Amount);
        }
        // Source exemption information is normally held on the header breakdown, not each line.
        foreach (InvoiceDeclaredTax declared in invoice.DeclaredTaxes) {
            if (declared == null || declared.Category == null) throw new ArgumentException("A declared VAT breakdown is null.", nameof(invoice));
            if (groups.TryGetValue((declared.Category.Code, NormalizeRate(declared.Category)), out TaxGroup? group)) {
                if (preserveDeclaredAmounts) {
                    group.MergeReason(declared.Category);
                    group.DeclaredAmount = declared.TaxAmount;
                } else group.UseSourceReasonWhenMissing(declared.Category);
            }
        }
        var taxes = groups.Values.OrderBy(group => group.Code, StringComparer.Ordinal).ThenBy(group => group.Rate)
            .Select(group => new InvoiceCalculatedTax(group.Code, group.Rate, group.Basis, group.Reason, group.ReasonCode, group.DeclaredAmount)).ToList();
        return new InvoiceCalculation(lines.AsReadOnly(), taxes.AsReadOnly(), allowances, charges, invoice.PrepaidAmount, invoice.RoundingAmount);
    }

    private static void CheckAmount(InvoiceAllowanceCharge? item) {
        if (item == null || item.Amount < 0m || RoundAmount(item.Amount) != item.Amount)
            throw new ArgumentException("Allowances and charges require a non-negative amount with at most two decimal places.");
    }

    private static void AddTax(Dictionary<(string Code, decimal? Rate), TaxGroup> groups, InvoiceTaxCategory? category, decimal basis) {
        if (category == null || string.IsNullOrWhiteSpace(category.Code) || category.Rate < 0m)
            throw new ArgumentException("Each line and document adjustment requires a VAT category and a non-negative rate when present.");
        var key = (category.Code, NormalizeRate(category));
        if (!groups.TryGetValue(key, out TaxGroup? group)) { group = new TaxGroup(category); groups.Add(key, group); }
        else group.MergeReason(category);
        group.Basis = InvoiceArithmetic.Add(group.Basis, basis);
    }

    internal static decimal? NormalizeRate(InvoiceTaxCategory category) => category.Code == "O" && category.Rate == 0m ? null : category.Rate;

    private sealed class TaxGroup {
        internal TaxGroup(InvoiceTaxCategory category) { Code = category.Code; Rate = NormalizeRate(category); MergeReason(category); }
        internal string Code { get; }
        internal decimal? Rate { get; }
        internal decimal Basis { get; set; }
        internal decimal? DeclaredAmount { get; set; }
        internal string? Reason { get; private set; }
        internal string? ReasonCode { get; private set; }
        internal void UseSourceReasonWhenMissing(InvoiceTaxCategory category) {
            // An edited exemption replaces the source pair; header-only imported details remain valid.
            if (Reason != null || ReasonCode != null) return;
            Reason = category.ExemptionReason;
            ReasonCode = category.ExemptionReasonCode;
        }
        internal void MergeReason(InvoiceTaxCategory category) {
            if (Reason != null && category.ExemptionReason != null && Reason != category.ExemptionReason ||
                ReasonCode != null && category.ExemptionReasonCode != null && ReasonCode != category.ExemptionReasonCode)
                throw new ArgumentException("A VAT category/rate has conflicting exemption reasons.");
            Reason = Reason ?? category.ExemptionReason;
            ReasonCode = ReasonCode ?? category.ExemptionReasonCode;
        }
    }
}
