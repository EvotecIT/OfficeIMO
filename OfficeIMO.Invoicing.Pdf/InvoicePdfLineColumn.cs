namespace OfficeIMO.Invoicing.Pdf;

/// <summary>Identifies an ordered column in the visible invoice line table.</summary>
public enum InvoicePdfLineColumn {
    /// <summary>Item name and business details not assigned to another column.</summary>
    Item,
    /// <summary>Invoiced quantity and unit.</summary>
    Quantity,
    /// <summary>Net price per price-base quantity.</summary>
    NetPrice,
    /// <summary>VAT category and rate.</summary>
    Vat,
    /// <summary>Line net amount.</summary>
    NetAmount,
    /// <summary>Unit of measure.</summary>
    Unit,
    /// <summary>Line identifier.</summary>
    LineIdentifier,
    /// <summary>Item description.</summary>
    Description,
    /// <summary>Service period.</summary>
    Period,
    /// <summary>Seller-assigned item identifier.</summary>
    SellerItem,
    /// <summary>Buyer-assigned item identifier.</summary>
    BuyerItem,
    /// <summary>Standard item identifier and its scheme.</summary>
    StandardItem,
    /// <summary>Buyer accounting reference.</summary>
    AccountingReference,
    /// <summary>Gross price per price-base quantity.</summary>
    GrossPrice,
    /// <summary>Price discount per price-base quantity.</summary>
    PriceDiscount
}

/// <summary>Controls presentation of machine-readable unit and payment codes.</summary>
public enum InvoicePdfCodeDisplay {
    /// <summary>Display the original code.</summary>
    Code,
    /// <summary>Display its translated description, falling back to the code when unknown.</summary>
    Description,
    /// <summary>Display the code followed by its translated description.</summary>
    CodeAndDescription
}
