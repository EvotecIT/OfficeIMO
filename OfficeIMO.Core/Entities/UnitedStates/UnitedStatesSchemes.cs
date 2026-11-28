namespace OfficeIMO.Entities;

/// <summary>United States Social Security and employer identification schemes.</summary>
public static class UnitedStatesSchemes {
    /// <summary>Social Security number.</summary>
    public static EntityScheme Ssn { get; } = new SsnScheme();
    /// <summary>Employer Identification Number.</summary>
    public static EntityScheme Ein { get; } = new EinScheme();
}
