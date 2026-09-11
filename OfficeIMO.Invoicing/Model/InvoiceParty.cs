namespace OfficeIMO.Invoicing;

/// <summary>Business identifier and its optional issuing scheme.</summary>
public sealed class InvoiceIdentifier {
    /// <summary>Creates an identifier.</summary>
    public InvoiceIdentifier(string value, string? schemeId = null) { Value = value; SchemeId = schemeId; }
    /// <summary>Identifier value.</summary>
    public string Value { get; set; }
    /// <summary>Identifier scheme, such as an ISO 6523 ICD or electronic address scheme code.</summary>
    public string? SchemeId { get; set; }
}

/// <summary>Postal address (BG-5/BG-8/BG-12/BG-15).</summary>
public sealed class InvoiceAddress {
    /// <summary>First address line.</summary>
    public string? Line1 { get; set; }
    /// <summary>Second address line.</summary>
    public string? Line2 { get; set; }
    /// <summary>Third address line.</summary>
    public string? Line3 { get; set; }
    /// <summary>City name.</summary>
    public string? City { get; set; }
    /// <summary>Post code.</summary>
    public string? PostCode { get; set; }
    /// <summary>Country subdivision.</summary>
    public string? Subdivision { get; set; }
    /// <summary>ISO 3166-1 alpha-2 country code.</summary>
    public string CountryCode { get; set; } = string.Empty;
}

/// <summary>Contact details (BG-6/BG-9).</summary>
public sealed class InvoiceContact {
    /// <summary>Contact point or person's name.</summary>
    public string? Name { get; set; }
    /// <summary>Telephone number.</summary>
    public string? Telephone { get; set; }
    /// <summary>Email address.</summary>
    public string? Email { get; set; }
}

/// <summary>Party identity, registration, address and contact information.</summary>
public sealed class InvoiceParty {
    /// <summary>Legal name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Trading name.</summary>
    public string? TradingName { get; set; }
    /// <summary>Additional legal information (BT-33 for a seller).</summary>
    public string? LegalInformation { get; set; }
    /// <summary>Business identifiers.</summary>
    public IList<InvoiceIdentifier> Identifiers { get; } = new List<InvoiceIdentifier>();
    /// <summary>Legal registration identifier.</summary>
    public InvoiceIdentifier? LegalRegistration { get; set; }
    /// <summary>VAT registration identifier including its country prefix.</summary>
    public string? VatIdentifier { get; set; }
    /// <summary>Other tax registration identifier.</summary>
    public string? TaxRegistration { get; set; }
    /// <summary>Electronic address with an explicit address scheme.</summary>
    public InvoiceIdentifier? ElectronicAddress { get; set; }
    /// <summary>Postal address.</summary>
    public InvoiceAddress Address { get; set; } = new InvoiceAddress();
    /// <summary>Contact information.</summary>
    public InvoiceContact? Contact { get; set; }
}
