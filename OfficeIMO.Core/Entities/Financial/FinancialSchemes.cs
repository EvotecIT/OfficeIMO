namespace OfficeIMO.Entities;

/// <summary>International bank, payment card, securities and legal entity schemes.</summary>
public static class FinancialSchemes {
    /// <summary>International Bank Account Number.</summary>
    public static EntityScheme Iban => IbanHolder.Instance;
    /// <summary>Business identifier code.</summary>
    public static EntityScheme Bic => Schemes.Bic;
    /// <summary>Payment card number.</summary>
    public static EntityScheme Card => Schemes.Card;
    /// <summary>International Securities Identification Number.</summary>
    public static EntityScheme Isin => Schemes.Isin;
    /// <summary>Legal Entity Identifier.</summary>
    public static EntityScheme Lei => Schemes.Lei;

    private static class IbanHolder {
        internal static readonly EntityScheme Instance = new IbanScheme();
    }

    private static class Schemes {
        internal static readonly EntityScheme Bic = new BicScheme();
        internal static readonly EntityScheme Card = new CardScheme();
        internal static readonly EntityScheme Isin = new IsinScheme();
        internal static readonly EntityScheme Lei = new LeiScheme();
    }
}
