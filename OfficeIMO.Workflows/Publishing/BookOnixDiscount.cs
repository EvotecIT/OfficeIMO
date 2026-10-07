namespace OfficeIMO.Workflows;

/// <summary>Business-to-business discount application (ONIX list 170).</summary>
public enum BookOnixDiscountKind {
    /// <summary>Applies to all units in a qualifying order (01).</summary>
    Rising,
    /// <summary>May apply retrospectively based on orders over an agreed period (02).</summary>
    RisingCumulative,
    /// <summary>Applies to marginal units in a qualifying order (03).</summary>
    Progressive,
    /// <summary>Counts previous orders over an agreed period when applying marginal discounts (04).</summary>
    ProgressiveCumulative
}

/// <summary>
/// Explicit business-to-business discount on the containing price. At least Percent or Amount is required.
/// Values are serialized without calculating a payable price, tier selection, stacking, or rounding.
/// </summary>
public sealed record BookOnixDiscount {
    /// <summary>How the discount applies. Cumulative periods are defined by trading-partner agreement.</summary>
    public required BookOnixDiscountKind Kind { get; init; }
    /// <summary>Optional percentage in the inclusive range 0–100.</summary>
    public decimal? Percent { get; init; }
    /// <summary>Optional nonnegative discount per copy, in the price's currency, at most the price amount.</summary>
    public decimal? Amount { get; init; }
    /// <summary>Optional positive minimum number of copies. An omitted value leaves the quantity unspecified.</summary>
    public int? MinimumQuantity { get; init; }
    /// <summary>Optional maximum copies, at least MinimumQuantity. Requires MinimumQuantity.</summary>
    public int? MaximumQuantity { get; init; }
}
