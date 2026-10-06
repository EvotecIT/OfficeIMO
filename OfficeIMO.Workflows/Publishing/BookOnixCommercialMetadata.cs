namespace OfficeIMO.Workflows;

/// <summary>Explicit ONIX publishing status (list 64), independent of each supplier's availability.</summary>
public enum BookOnixPublishingStatus {
    /// <summary>Announced product abandoned (01); no publication date may be supplied.</summary>
    Cancelled,
    /// <summary>Not yet published (02); an expected publication date is required.</summary>
    Forthcoming,
    /// <summary>Postponed without a known date (03); no publication date may be supplied.</summary>
    PostponedIndefinitely,
    /// <summary>Published and active (04).</summary>
    Active,
    /// <summary>Permanently inactive at the publisher (07).</summary>
    OutOfPrint,
    /// <summary>Withdrawn from sale (11).</summary>
    Withdrawn
}

/// <summary>Publisher-supplied sales rights for a territory (ONIX list 46).</summary>
public enum BookOnixSalesRightsKind {
    /// <summary>For sale with exclusive publishing rights (01).</summary>
    Exclusive,
    /// <summary>For sale with non-exclusive publishing rights (02).</summary>
    NonExclusive,
    /// <summary>Not for sale, reason unspecified (03).</summary>
    NotForSale,
    /// <summary>Not for sale despite exclusive rights (04).</summary>
    NotForSaleExclusive,
    /// <summary>Not for sale despite non-exclusive rights (05).</summary>
    NotForSaleNonExclusive,
    /// <summary>Not for sale because publishing rights are not held (06).</summary>
    RightsNotHeld
}

/// <summary>An explicit country territory or WORLD with optional country exclusions.</summary>
public sealed record BookOnixTerritory {
    /// <summary>Uppercase ONIX list 91 country codes; mutually exclusive with Worldwide.</summary>
    public IReadOnlyList<string> Countries { get; init; } = [];
    /// <summary>Explicit worldwide scope; false does not imply any default territory.</summary>
    public bool Worldwide { get; init; }
    /// <summary>Country exclusions, allowed only with Worldwide. Each country list is bounded to 250 unique codes.</summary>
    public IReadOnlyList<string> ExcludedCountries { get; init; } = [];
}

/// <summary>A publisher's territorial sales-rights assertion; OfficeIMO does not verify ownership.</summary>
public sealed record BookOnixSalesRights(BookOnixSalesRightsKind Kind, BookOnixTerritory Territory);

/// <summary>Explicit ONIX supplier roles supported by digital-book delivery.</summary>
public enum BookOnixSupplierRole {
    /// <summary>Publisher supplying resellers (01).</summary>
    PublisherToResellers,
    /// <summary>Exclusive distributor supplying resellers (02).</summary>
    ExclusiveDistributorToResellers,
    /// <summary>Non-exclusive distributor supplying resellers (03).</summary>
    NonExclusiveDistributorToResellers,
    /// <summary>Wholesaler supplying retailers (04).</summary>
    Wholesaler,
    /// <summary>Retailer (08).</summary>
    Retailer,
    /// <summary>Publisher supplying end customers (09).</summary>
    PublisherToCustomers,
    /// <summary>Exclusive distributor supplying end customers (10).</summary>
    ExclusiveDistributorToCustomers,
    /// <summary>Non-exclusive distributor supplying end customers (11).</summary>
    NonExclusiveDistributorToCustomers
}

/// <summary>Supplier-specific availability (ONIX list 65).</summary>
public enum BookOnixAvailability {
    /// <summary>Not yet available (10).</summary>
    NotYetAvailable,
    /// <summary>Available (20).</summary>
    Available,
    /// <summary>Temporarily unavailable (30).</summary>
    TemporarilyUnavailable,
    /// <summary>Not available, reason unspecified (40).</summary>
    Unavailable,
    /// <summary>Withdrawn from sale by this supplier (46).</summary>
    Withdrawn
}

/// <summary>Explicit reason for omitting a price (ONIX list 57).</summary>
public enum BookOnixUnpricedKind {
    /// <summary>Free of charge (01).</summary>
    Free,
    /// <summary>Price not yet announced (02).</summary>
    ToBeAnnounced,
    /// <summary>Contact the supplier for pricing (04).</summary>
    ContactSupplier
}

/// <summary>Price basis and tax inclusion (ONIX list 58). No tax amount is inferred or calculated.</summary>
public enum BookOnixPriceKind {
    /// <summary>Recommended retail price excluding tax (01).</summary>
    RecommendedExcludingTax,
    /// <summary>Recommended retail price including tax (02).</summary>
    RecommendedIncludingTax,
    /// <summary>Fixed retail price excluding tax (03).</summary>
    FixedExcludingTax,
    /// <summary>Fixed retail price including tax (04).</summary>
    FixedIncludingTax,
    /// <summary>Supplier's net price excluding tax (05).</summary>
    SupplierNetExcludingTax,
    /// <summary>Publisher's agency retail price excluding tax (41).</summary>
    PublisherRetailExcludingTax,
    /// <summary>Publisher's agency retail price including tax (42).</summary>
    PublisherRetailIncludingTax
}

/// <summary>A positive monetary price for a specified market. Decimal values are serialized without rounding.</summary>
public sealed record BookOnixPrice {
    /// <summary>Explicit price basis.</summary>
    public required BookOnixPriceKind Kind { get; init; }
    /// <summary>Positive amount. Use Unpriced = Free instead of a zero price.</summary>
    public required decimal Amount { get; init; }
    /// <summary>Uppercase three-letter ONIX list 96 currency code.</summary>
    public required string CurrencyCode { get; init; }
    /// <summary>At most 16 explicit business-to-business discounts, retained in supplied order without tier selection or stacking.</summary>
    public IReadOnlyList<BookOnixDiscount> Discounts { get; init; } = [];
    /// <summary>At most 16 explicit tax components, only for tax-inclusive price kinds. Mutually exclusive with TaxExempt.</summary>
    public IReadOnlyList<BookOnixTax> Taxes { get; init; } = [];
    /// <summary>Explicit tax exemption. An empty Taxes list does not imply exemption or zero rating.</summary>
    public bool TaxExempt { get; init; }
    /// <summary>Optional narrower price territory; when omitted, uses the supply's explicit market.</summary>
    public BookOnixTerritory? Territory { get; init; }
    /// <summary>Date on which the price becomes effective (price-date role 14).</summary>
    public DateOnly? ValidFrom { get; init; }
    /// <summary>Date on which the price ceases to be effective (price-date role 15).</summary>
    public DateOnly? ValidUntil { get; init; }
}

/// <summary>One explicit market, supplier and availability declaration, with prices or an unpriced reason.</summary>
public sealed record BookOnixSupply {
    /// <summary>Market to which this supply declaration applies.</summary>
    public required BookOnixTerritory Territory { get; init; }
    /// <summary>Supplier's name.</summary>
    public required string SupplierName { get; init; }
    /// <summary>Supplier's role; not inferred from PublisherName.</summary>
    public required BookOnixSupplierRole SupplierRole { get; init; }
    /// <summary>Availability from this supplier.</summary>
    public required BookOnixAvailability Availability { get; init; }
    /// <summary>Expected availability date (supply-date role 08) for forthcoming or temporary unavailability.</summary>
    public DateOnly? ExpectedSupplyDate { get; init; }
    /// <summary>Explicit exception when no expected availability date is known; mutually exclusive with a date.</summary>
    public bool ExpectedSupplyDateUnknown { get; init; }
    /// <summary>At most 16 prices. Mutually exclusive with Unpriced.</summary>
    public IReadOnlyList<BookOnixPrice> Prices { get; init; } = [];
    /// <summary>Explicit reason for no price. Missing prices never imply free supply.</summary>
    public BookOnixUnpricedKind? Unpriced { get; init; }
}

/// <summary>Publisher-supplied commercial assertions. No currency conversion, tax calculation or rights lookup is performed.</summary>
public sealed record BookOnixCommercialMetadata {
    /// <summary>Optional product-level publishing status.</summary>
    public BookOnixPublishingStatus? PublishingStatus { get; init; }
    /// <summary>At most 32 non-overlapping rights territories. Undeclared territories remain unstated.</summary>
    public IReadOnlyList<BookOnixSalesRights> SalesRights { get; init; } = [];
    /// <summary>At most 32 supply declarations. Available or expected supply requires declared for-sale rights.</summary>
    public IReadOnlyList<BookOnixSupply> Supplies { get; init; } = [];
}
