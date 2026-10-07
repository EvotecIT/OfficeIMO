namespace OfficeIMO.Workflows;

/// <summary>Explicit tax type (ONIX list 171).</summary>
public enum BookOnixTaxType {
    /// <summary>Value-added tax, including GST (01).</summary>
    ValueAdded,
    /// <summary>Retail sales tax (02).</summary>
    Sales,
    /// <summary>Separately identified environmental tax (03).</summary>
    Environmental
}

/// <summary>Publisher-supplied tax-rate classification (ONIX list 62).</summary>
public enum BookOnixTaxRateCode {
    /// <summary>Higher than standard (H).</summary>
    Higher,
    /// <summary>Tax paid at source (P).</summary>
    PaidAtSource,
    /// <summary>Lower rate (R).</summary>
    Lower,
    /// <summary>Standard rate (S).</summary>
    Standard,
    /// <summary>Super-low rate (T).</summary>
    SuperLow,
    /// <summary>Zero-rated (Z), distinct from tax exemption.</summary>
    Zero
}

/// <summary>
/// A tax component included in an ONIX price. Values are explicit assertions, serialized without rounding.
/// At least RatePercent or Amount is required. Currency and territory come from the containing price.
/// </summary>
public sealed record BookOnixTax {
    /// <summary>Type of tax included in the price.</summary>
    public required BookOnixTaxType Type { get; init; }
    /// <summary>Optional tax-rate classification; OfficeIMO does not select a jurisdictional rate.</summary>
    public BookOnixTaxRateCode? RateCode { get; init; }
    /// <summary>Optional percentage between zero and 100 inclusive.</summary>
    public decimal? RatePercent { get; init; }
    /// <summary>Optional positive part of the price before this tax, at most the containing price amount.</summary>
    public decimal? TaxableAmount { get; init; }
    /// <summary>Optional nonnegative tax amount included in the price. Supplied tax amounts cannot sum above the price.</summary>
    public decimal? Amount { get; init; }
    /// <summary>Optional description of the price component subject to this tax, at most 4096 characters.</summary>
    public string? PricePartDescription { get; init; }
}
