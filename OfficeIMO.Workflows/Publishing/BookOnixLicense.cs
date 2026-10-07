namespace OfficeIMO.Workflows;

/// <summary>A plain-text license name, optionally tagged with an ONIX list 74 language.</summary>
public sealed record BookOnixLicenseName(string Text, string? LanguageCode = null);

/// <summary>ONIX list 218 license expression formats.</summary>
public enum BookOnixLicenseExpressionType {
    /// <summary>Terms intended for the general reader (01).</summary>
    HumanReadable,
    /// <summary>Terms intended for a legal specialist (02).</summary>
    ProfessionalReadable,
    /// <summary>A separately obtainable additional license for the general reader (03).</summary>
    AdditionalHumanReadable,
    /// <summary>A separately obtainable additional license for a legal specialist (04).</summary>
    AdditionalProfessionalReadable,
    /// <summary>ONIX-PL expression (10).</summary>
    OnixPl,
    /// <summary>ODRL expression (20).</summary>
    Odrl,
    /// <summary>ODRL expression for a separately obtainable additional license (21).</summary>
    AdditionalOdrl
}

/// <summary>An explicit expression format and absolute HTTP(S) link without credentials. Links are not fetched.</summary>
public sealed record BookOnixLicenseExpression(BookOnixLicenseExpressionType Type, string Link);

/// <summary>Publisher-supplied license metadata. No rights are inferred or enforced.</summary>
public sealed record BookOnixLicense {
    /// <summary>One to 16 plain-text names, at most 100 UTF-16 code units each. Repeated names require distinct explicit languages.</summary>
    public required IReadOnlyList<BookOnixLicenseName> Names { get; init; }
    /// <summary>Up to 16 distinct format/link pairs. Each link uses the existing 4096-character ONIX field bound.</summary>
    public IReadOnlyList<BookOnixLicenseExpression> Expressions { get; init; } = [];
    /// <summary>Optional inclusive first effective date (role 14).</summary>
    public DateOnly? ValidFrom { get; init; }
    /// <summary>Optional inclusive last effective date (role 15). Must not precede ValidFrom.</summary>
    public DateOnly? ValidUntil { get; init; }
}
