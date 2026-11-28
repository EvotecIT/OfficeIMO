namespace OfficeIMO.Entities;

/// <summary>United Kingdom national, company and VAT schemes.</summary>
public static class UnitedKingdomSchemes {
    /// <summary>National Insurance number.</summary>
    public static EntityScheme Nino { get; } = new NinoScheme();
    /// <summary>Companies House registration number.</summary>
    public static EntityScheme Company { get; } = new CompaniesHouseScheme();
    /// <summary>United Kingdom VAT registration number.</summary>
    public static EntityScheme Vat { get; } = new UkVatScheme("GB", "GB.VAT");
}
