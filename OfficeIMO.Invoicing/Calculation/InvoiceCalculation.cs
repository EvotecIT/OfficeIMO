namespace OfficeIMO.Invoicing;

/// <summary>Immutable calculated monetary results in the invoice currency.</summary>
public sealed class InvoiceCalculation {
    internal InvoiceCalculation(IReadOnlyList<InvoiceCalculatedLine> lines, IReadOnlyList<InvoiceCalculatedTax> taxes,
        decimal allowances, decimal charges, decimal prepaid, decimal rounding) {
        Lines = lines;
        Taxes = taxes;
        LineNetTotal = lines.Sum(line => line.NetAmount);
        AllowanceTotal = allowances;
        ChargeTotal = charges;
        TaxExclusiveTotal = LineNetTotal - allowances + charges;
        TaxTotal = taxes.Sum(tax => tax.TaxAmount);
        TaxInclusiveTotal = TaxExclusiveTotal + TaxTotal;
        PrepaidAmount = prepaid;
        RoundingAmount = rounding;
        PayableAmount = TaxInclusiveTotal - prepaid + rounding;
    }

    /// <summary>Calculated lines in source order.</summary>
    public IReadOnlyList<InvoiceCalculatedLine> Lines { get; }
    /// <summary>VAT breakdowns in stable category/rate order.</summary>
    public IReadOnlyList<InvoiceCalculatedTax> Taxes { get; }
    /// <summary>Sum of line net amounts (BT-106).</summary>
    public decimal LineNetTotal { get; }
    /// <summary>Document allowances (BT-107).</summary>
    public decimal AllowanceTotal { get; }
    /// <summary>Document charges (BT-108).</summary>
    public decimal ChargeTotal { get; }
    /// <summary>Total excluding VAT (BT-109).</summary>
    public decimal TaxExclusiveTotal { get; }
    /// <summary>Total VAT (BT-110).</summary>
    public decimal TaxTotal { get; }
    /// <summary>Total including VAT (BT-112).</summary>
    public decimal TaxInclusiveTotal { get; }
    /// <summary>Prepaid amount (BT-113).</summary>
    public decimal PrepaidAmount { get; }
    /// <summary>Rounding adjustment (BT-114).</summary>
    public decimal RoundingAmount { get; }
    /// <summary>Amount due (BT-115).</summary>
    public decimal PayableAmount { get; }
}

/// <summary>Immutable line calculation, indexed in the same order as the source model.</summary>
public sealed class InvoiceCalculatedLine {
    internal InvoiceCalculatedLine(string id, decimal formulaNetAmount, decimal? declared) { Id = id; FormulaNetAmount = formulaNetAmount; NetAmount = declared ?? InvoiceCalculator.RoundAmount(formulaNetAmount); }
    /// <summary>Source line identifier.</summary>
    public string Id { get; }
    /// <summary>Unrounded quantity, price and adjustment formula before source declarations.</summary>
    public decimal FormulaNetAmount { get; }
    /// <summary>Source-declared line net amount (BT-131), or the formula amount when absent.</summary>
    public decimal NetAmount { get; }
}

/// <summary>Immutable calculated VAT breakdown.</summary>
public sealed class InvoiceCalculatedTax {
    internal InvoiceCalculatedTax(string code, decimal? rate, decimal basis, string? reason, string? reasonCode, decimal? declaredAmount) {
        CategoryCode = code; Rate = rate; TaxableAmount = basis;
        ExemptionReason = reason; ExemptionReasonCode = reasonCode;
        FormulaTaxAmount = InvoiceCalculator.RoundAmount(basis * (rate ?? 0m) / 100m);
        TaxAmount = declaredAmount ?? FormulaTaxAmount;
    }
    /// <summary>VAT category code.</summary>
    public string CategoryCode { get; }
    /// <summary>VAT percentage.</summary>
    public decimal? Rate { get; }
    /// <summary>Taxable amount (BT-116).</summary>
    public decimal TaxableAmount { get; }
    /// <summary>Rounded taxable basis multiplied by the VAT percentage.</summary>
    public decimal FormulaTaxAmount { get; }
    /// <summary>Source-declared VAT amount (BT-117), or the formula amount when absent.</summary>
    public decimal TaxAmount { get; }
    /// <summary>Exemption reason (BT-120).</summary>
    public string? ExemptionReason { get; }
    /// <summary>Exemption reason code (BT-121).</summary>
    public string? ExemptionReasonCode { get; }
}
