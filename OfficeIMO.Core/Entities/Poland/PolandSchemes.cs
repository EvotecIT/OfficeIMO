namespace OfficeIMO.Entities;

/// <summary>Polish personal, tax, company and identity document schemes.</summary>
public static class PolandSchemes {
    /// <summary>Population register number.</summary>
    public static EntityScheme Pesel => Schemes.Pesel;
    /// <summary>Tax identification number.</summary>
    public static EntityScheme Nip => Schemes.Nip;
    /// <summary>Statistical business registry number.</summary>
    public static EntityScheme Regon => Schemes.Regon;
    /// <summary>National Court Register number.</summary>
    public static EntityScheme Krs => Schemes.Krs;
    /// <summary>Identity card number.</summary>
    public static EntityScheme IdentityCard => Schemes.IdentityCard;
    /// <summary>Passport number.</summary>
    public static EntityScheme Passport => Schemes.Passport;

    private static class Schemes {
        internal static readonly EntityScheme Pesel = new PeselScheme();
        internal static readonly EntityScheme Nip = new NipScheme();
        internal static readonly EntityScheme Regon = new RegonScheme();
        internal static readonly EntityScheme Krs = new KrsScheme();
        internal static readonly EntityScheme IdentityCard = new IdentityCardScheme();
        internal static readonly EntityScheme Passport = new PassportScheme();
    }
}
