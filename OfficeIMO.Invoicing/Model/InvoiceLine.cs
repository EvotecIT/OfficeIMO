namespace OfficeIMO.Invoicing;

/// <summary>VAT category, rate and exemption information.</summary>
public sealed class InvoiceTaxCategory {
    /// <summary>UNCL 5305 category: S, Z, E, AE, K, G, O, L or M.</summary>
    public string Code { get; set; } = "S";
    /// <summary>VAT percentage; absent for a category that prohibits a rate.</summary>
    public decimal? Rate { get; set; }
    /// <summary>Exemption reason text (BT-120).</summary>
    public string? ExemptionReason { get; set; }
    /// <summary>VATEX exemption reason code (BT-121).</summary>
    public string? ExemptionReasonCode { get; set; }
}

/// <summary>Line or document-level allowance or charge.</summary>
public sealed class InvoiceAllowanceCharge {
    /// <summary>True for a charge; false for an allowance.</summary>
    public bool IsCharge { get; set; }
    /// <summary>Positive monetary amount.</summary>
    public decimal Amount { get; set; }
    /// <summary>Base amount used with a percentage.</summary>
    public decimal? BaseAmount { get; set; }
    /// <summary>Percentage, expressed as 10 for ten percent.</summary>
    public decimal? Percentage { get; set; }
    /// <summary>Reason text.</summary>
    public string? Reason { get; set; }
    /// <summary>UNTDID 5189 allowance or 7161 charge reason code.</summary>
    public string? ReasonCode { get; set; }
    /// <summary>Required for document-level amounts; line amounts inherit the line's category.</summary>
    public InvoiceTaxCategory? Tax { get; set; }
}

/// <summary>Product classification identifier and its list metadata.</summary>
public sealed class InvoiceItemClassification {
    /// <summary>Classification value (BT-158).</summary>
    public string Value { get; set; } = string.Empty;
    /// <summary>UNCL 7143 list identifier.</summary>
    public string ListId { get; set; } = string.Empty;
    /// <summary>List version.</summary>
    public string? ListVersion { get; set; }
}

/// <summary>Item attribute (BG-32).</summary>
public sealed class InvoiceItemAttribute {
    /// <summary>Attribute name (BT-160).</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Attribute value (BT-161).</summary>
    public string Value { get; set; } = string.Empty;
}

/// <summary>Invoice line including item, quantity, price, adjustments and VAT.</summary>
public sealed class InvoiceLine {
    /// <summary>Line identifier (BT-126).</summary>
    public string Id { get; set; } = string.Empty;
    /// <summary>Item name (BT-153).</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Item description (BT-154).</summary>
    public string? Description { get; set; }
    /// <summary>Line note (BT-127).</summary>
    public string? Note { get; set; }
    /// <summary>Invoiced quantity (BT-129).</summary>
    public decimal Quantity { get; set; }
    /// <summary>UN/ECE recommendation 20/21 unit code (BT-130).</summary>
    public string UnitCode { get; set; } = "C62";
    /// <summary>Item net price per base quantity (BT-146).</summary>
    public decimal UnitPrice { get; set; }
    /// <summary>Quantity to which the price applies (BT-149).</summary>
    public decimal PriceBaseQuantity { get; set; } = 1m;
    /// <summary>Item gross price before the price discount (BT-148).</summary>
    public decimal? GrossPrice { get; set; }
    /// <summary>Discount per base quantity (BT-147).</summary>
    public decimal? PriceDiscount { get; set; }
    /// <summary>VAT category and percentage (BG-30).</summary>
    public InvoiceTaxCategory Tax { get; set; } = new InvoiceTaxCategory();
    /// <summary>Line allowances and charges (BG-27/BG-28).</summary>
    public IList<InvoiceAllowanceCharge> AllowancesAndCharges { get; } = new List<InvoiceAllowanceCharge>();
    /// <summary>Line period (BG-26).</summary>
    public InvoicePeriod? Period { get; set; }
    /// <summary>Purchase order line reference (BT-132).</summary>
    public string? OrderLineReference { get; set; }
    /// <summary>Buyer accounting reference (BT-133).</summary>
    public string? AccountingReference { get; set; }
    /// <summary>Invoiced object identifier (BT-128).</summary>
    public InvoiceIdentifier? ObjectIdentifier { get; set; }
    /// <summary>Seller item identifier (BT-155).</summary>
    public string? SellerItemIdentifier { get; set; }
    /// <summary>Buyer item identifier (BT-156).</summary>
    public string? BuyerItemIdentifier { get; set; }
    /// <summary>Standard item identifier with scheme (BT-157).</summary>
    public InvoiceIdentifier? StandardItemIdentifier { get; set; }
    /// <summary>Country of origin (BT-159).</summary>
    public string? OriginCountryCode { get; set; }
    /// <summary>Item classifications (BT-158).</summary>
    public IList<InvoiceItemClassification> Classifications { get; } = new List<InvoiceItemClassification>();
    /// <summary>Additional item attributes (BG-32).</summary>
    public IList<InvoiceItemAttribute> Attributes { get; } = new List<InvoiceItemAttribute>();
    /// <summary>Source-declared net line amount (BT-131), kept for comparison with the calculation.</summary>
    public decimal? DeclaredNetAmount { get; set; }
}
