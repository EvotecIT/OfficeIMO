namespace OfficeIMO.Workflows;

/// <summary>ONIX list 100 discount or commission code scheme.</summary>
public enum BookOnixDiscountScheme {
    /// <summary>BIC discount group (01).</summary>
    BicDiscount,
    /// <summary>Named proprietary trade-discount scheme (02).</summary>
    ProprietaryDiscount,
    /// <summary>Dutch Boeksoort terms (03).</summary>
    Boeksoort,
    /// <summary>German terms (04).</summary>
    GermanTerms,
    /// <summary>Named proprietary agency-commission scheme (05).</summary>
    ProprietaryCommission,
    /// <summary>BIC commission group (06).</summary>
    BicCommission,
    /// <summary>ISNI-based discount group (07).</summary>
    IsniDiscount
}

/// <summary>An explicit trading-partner code; its monetary meaning is not inferred or looked up.</summary>
public sealed record BookOnixDiscountCode {
    /// <summary>Scheme defining the code's meaning.</summary>
    public required BookOnixDiscountScheme Scheme { get; init; }
    /// <summary>Code retained verbatim, at most 4096 characters. BIC and ISNI-based codes require their structural formats.</summary>
    public required string Code { get; init; }
    /// <summary>Distinctive scheme name, required for proprietary discount/commission schemes and forbidden for other schemes.</summary>
    public string? SchemeName { get; init; }
}
