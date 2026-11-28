namespace OfficeIMO.Entities;

internal static class PolandKeywords {
    internal static void Register(EntityCatalogBuilder builder) {
        builder.AddKeyword("PL.PESEL", "pl", "PESEL").AddKeyword("PL.PESEL", "en", "PESEL")
            .AddKeyword("PL.PESEL", "pl", "numer ewidencyjny")
            .AddKeyword("PL.NIP", "pl", "NIP").AddKeyword("PL.NIP", "pl", "numer identyfikacji podatkowej")
            .AddKeyword("PL.NIP", "en", "tax identification number")
            .AddKeyword("PL.REGON", "pl", "REGON")
            .AddKeyword("PL.KRS", "pl", "KRS").AddKeyword("PL.KRS", "pl", "Krajowy Rejestr Sądowy")
            .AddKeyword("PL.IDCARD", "pl", "nr dowodu osobistego").AddKeyword("PL.IDCARD", "pl", "dowód osobisty")
            .AddKeyword("PL.IDCARD", "pl", "numer dowodu").AddKeyword("PL.IDCARD", "pl", "seria i numer dowodu")
            .AddKeyword("PL.IDCARD", "en", "Polish identity card")
            .AddKeyword("PL.PASSPORT", "pl", "paszport").AddKeyword("PL.PASSPORT", "pl", "numer paszportu")
            .AddKeyword("PL.PASSPORT", "en", "Polish passport");
    }
}
