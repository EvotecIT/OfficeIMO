using System;

namespace OfficeIMO.Entities;

/// <summary>Registers the built-in regional and international recognition packs.</summary>
public static class EntityCatalogBuilderExtensions {
    /// <summary>Registers Polish national, tax, company and identity document identifiers.</summary>
    public static EntityCatalogBuilder AddPoland(this EntityCatalogBuilder builder) {
        AddPack(builder, "poland", EntityTokenShape.DigitRun | EntityTokenShape.AlphaNumericRun, PolandRecognizers.Recognize,
            PolandSchemes.Pesel, PolandSchemes.Nip, PolandSchemes.Regon, PolandSchemes.Krs, PolandSchemes.IdentityCard, PolandSchemes.Passport);

        PolandKeywords.Register(builder);
        return builder;
    }

    /// <summary>Registers VAT identifiers for VIES jurisdictions.</summary>
    public static EntityCatalogBuilder AddEuVat(this EntityCatalogBuilder builder) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        var schemes = new EntityScheme[EuVatSchemes.All.Count];
        for (int index = 0; index < schemes.Length; index++) schemes[index] = EuVatSchemes.All[index];
        AddPack(builder, "eu-vat", EntityTokenShape.DigitRun | EntityTokenShape.AlphaNumericRun, EuVatRecognizer.Recognize, schemes);
        EuVatKeywords.Register(builder);
        return builder;
    }

    /// <summary>Registers UK national insurance, company and VAT identifiers.</summary>
    public static EntityCatalogBuilder AddUnitedKingdom(this EntityCatalogBuilder builder) {
        AddPack(builder, "united-kingdom", EntityTokenShape.DigitRun | EntityTokenShape.AlphaNumericRun, UnitedKingdomRecognizers.Recognize,
            UnitedKingdomSchemes.Nino, UnitedKingdomSchemes.Company, UnitedKingdomSchemes.Vat);

        return builder.AddKeyword("GB.NINO", "en", "NINO").AddKeyword("GB.NINO", "en", "National Insurance number")
            .AddKeyword("GB.COMPANY", "en", "Companies House").AddKeyword("GB.COMPANY", "en", "company number")
            .AddKeyword("GB.COMPANY", "en", "company registration number").AddKeyword("GB.COMPANY", "en", "CRN")
            .AddKeyword("GB.VAT", "en", "VAT").AddKeyword("GB.VAT", "en", "VAT registration number");
    }

    /// <summary>Registers US Social Security and Employer Identification numbers.</summary>
    public static EntityCatalogBuilder AddUnitedStates(this EntityCatalogBuilder builder) {
        AddPack(builder, "united-states", EntityTokenShape.DigitRun, UnitedStatesRecognizers.Recognize,
            UnitedStatesSchemes.Ssn, UnitedStatesSchemes.Ein);

        return builder.AddKeyword("US.SSN", "en", "SSN").AddKeyword("US.SSN", "en", "social security number")
            .AddKeyword("US.EIN", "en", "EIN").AddKeyword("US.EIN", "en", "employer identification number")
            .AddKeyword("US.EIN", "en", "federal tax ID");
    }

    /// <summary>Registers IBAN, BIC, payment cards, ISIN and LEI identifiers.</summary>
    public static EntityCatalogBuilder AddFinancial(this EntityCatalogBuilder builder) {
        AddPack(builder, "financial", EntityTokenShape.DigitRun | EntityTokenShape.AlphaNumericRun, FinancialRecognizers.Recognize,
            FinancialSchemes.Iban, FinancialSchemes.Bic, FinancialSchemes.Card, FinancialSchemes.Isin, FinancialSchemes.Lei);

        return builder.AddKeyword("IBAN", "en", "IBAN").AddKeyword("IBAN", "pl", "nr rachunku")
            .AddKeyword("IBAN", "pl", "numer konta").AddKeyword("IBAN", "en", "account number").AddKeyword("IBAN", "de", "Kontonummer")
            .AddKeyword("BIC", "en", "BIC").AddKeyword("BIC", "en", "SWIFT").AddKeyword("BIC", "pl", "kod banku")
            .AddKeyword("CARD", "en", "card").AddKeyword("CARD", "pl", "karta").AddKeyword("CARD", "pl", "numer karty")
            .AddKeyword("ISIN", "en", "ISIN").AddKeyword("LEI", "en", "LEI");
    }

    /// <summary>Registers email addresses and URLs.</summary>
    public static EntityCatalogBuilder AddInternet(this EntityCatalogBuilder builder) {
        AddPack(builder, "internet", EntityTokenShape.At | EntityTokenShape.Scheme, InternetRecognizers.Recognize,
            InternetSchemes.Email, InternetSchemes.Url);

        return builder.AddKeyword("EMAIL", "en", "email").AddKeyword("EMAIL", "en", "e-mail")
            .AddKeyword("EMAIL", "pl", "adres e-mail").AddKeyword("URL", "en", "URL");
    }

    private static void AddPack(EntityCatalogBuilder builder, string id, EntityTokenShape shapes,
        Action<EntityRecognitionContext, EntityToken> recognize, params EntityScheme[] schemes) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        foreach (EntityScheme scheme in schemes) builder.Add(scheme);
        builder.AddRecognizer(new Internal.TokenEntityRecognizer(id, shapes, recognize, schemes));
    }
}
