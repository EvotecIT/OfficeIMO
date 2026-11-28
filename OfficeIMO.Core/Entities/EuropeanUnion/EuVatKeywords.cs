namespace OfficeIMO.Entities;

internal static class EuVatKeywords {
    internal static void Register(EntityCatalogBuilder builder) {
        // Use local tax names to recognize unprefixed national identifiers in context.
        Register(builder, "AT", "de", "UID", "USt-IdNr");
        Register(builder, "BE", "nl", "btw-nummer");
        Register(builder, "BE", "fr", "numéro de TVA belge");
        Register(builder, "BG", "bg", "ДДС", "номер по ДДС");
        Register(builder, "CY", "el", "αριθμός ΦΠΑ");
        Register(builder, "CZ", "cs", "DIČ", "DPH");
        Register(builder, "DE", "de", "USt-IdNr", "Umsatzsteuer-Identifikationsnummer");
        Register(builder, "DK", "da", "momsnummer", "CVR-nummer");
        Register(builder, "EE", "et", "käibemaksukohustuslase number", "KMKR");
        Register(builder, "EL", "el", "ΑΦΜ", "αριθμός φορολογικού μητρώου");
        Register(builder, "ES", "es", "NIF-IVA", "IVA intracomunitario");
        Register(builder, "FI", "fi", "ALV-tunnus", "arvonlisäverotunniste");
        Register(builder, "FI", "sv", "momsregistreringsnummer");
        Register(builder, "FR", "fr", "TVA intracommunautaire", "numéro de TVA");
        Register(builder, "HR", "hr", "PDV-ID", "PDV identifikacijski broj");
        Register(builder, "HU", "hu", "közösségi adószám", "adószám");
        Register(builder, "IE", "ga", "CBL", "uimhir CBL");
        Register(builder, "IE", "en", "Irish VAT");
        Register(builder, "IT", "it", "partita IVA", "P. IVA");
        Register(builder, "LT", "lt", "PVM kodas", "PVM mokėtojo kodas");
        Register(builder, "LU", "fr", "TVA luxembourgeoise", "numéro de TVA");
        Register(builder, "LV", "lv", "PVN numurs", "PVN reģistrācijas numurs");
        Register(builder, "MT", "mt", "numru tal-VAT");
        Register(builder, "MT", "en", "Maltese VAT");
        Register(builder, "NL", "nl", "btw-identificatienummer", "btw-id", "btw-nummer");
        Register(builder, "PL", "pl", "NIP-UE", "VAT-UE");
        Register(builder, "PT", "pt", "NIF-IVA", "número de IVA");
        Register(builder, "RO", "ro", "cod TVA", "cod de înregistrare în scopuri de TVA");
        Register(builder, "SE", "sv", "momsregistreringsnummer", "momsnummer");
        Register(builder, "SI", "sl", "ID za DDV", "identifikacijska številka za DDV");
        Register(builder, "SK", "sk", "IČ DPH", "identifikačné číslo pre DPH");
        Register(builder, "XI", "en", "Northern Ireland VAT");
    }

    private static void Register(EntityCatalogBuilder builder, string prefix, string language, params string[] terms) {
        foreach (string term in terms) builder.AddKeyword("EU.VAT." + prefix, language, term);
    }
}
