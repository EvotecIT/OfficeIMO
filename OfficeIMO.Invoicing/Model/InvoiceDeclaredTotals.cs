namespace OfficeIMO.Invoicing;

/// <summary>Source-declared invoice totals. Null means the field was absent, not zero.</summary>
public sealed class InvoiceDeclaredTotals {
    /// <summary>Sum of line net amounts (BT-106).</summary>
    public decimal? LineNetTotal { get; set; }
    /// <summary>Sum of document allowances (BT-107).</summary>
    public decimal? AllowanceTotal { get; set; }
    /// <summary>Sum of document charges (BT-108).</summary>
    public decimal? ChargeTotal { get; set; }
    /// <summary>Total excluding VAT (BT-109).</summary>
    public decimal? TaxExclusiveTotal { get; set; }
    /// <summary>Total VAT (BT-110).</summary>
    public decimal? TaxTotal { get; set; }
    /// <summary>Total including VAT (BT-112).</summary>
    public decimal? TaxInclusiveTotal { get; set; }
    /// <summary>Amount due (BT-115).</summary>
    public decimal? PayableAmount { get; set; }
}

/// <summary>Source-declared VAT breakdown.</summary>
public sealed class InvoiceDeclaredTax {
    /// <summary>VAT category, rate and exemption information.</summary>
    public InvoiceTaxCategory Category { get; set; } = new InvoiceTaxCategory();
    /// <summary>Taxable amount (BT-116).</summary>
    public decimal TaxableAmount { get; set; }
    /// <summary>Tax amount (BT-117).</summary>
    public decimal TaxAmount { get; set; }
}
