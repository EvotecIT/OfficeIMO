using System.Collections.ObjectModel;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class InvoicePdfLanguagePack {
    private IReadOnlyDictionary<string, string>? _unitDescriptions;
    private IReadOnlyDictionary<string, string>? _paymentDescriptions;

    /// <summary>Returns the translated description of a supported UN/ECE unit code, or null when unknown.</summary>
    public string? GetUnitDescription(string code) => CodeLabel(code, _unitDescriptions, UnitDescriptions);

    /// <summary>Formats a built-in unit in a quantity phrase with the plural form of the pack's language, retaining authored labels literally.</summary>
    internal string? GetQuantityUnitDescription(string code, decimal quantity) {
        string? label = GetUnitDescription(code);
        if (label == null || _unitDescriptions?.ContainsKey(code) == true || !UnitDescriptions.ContainsKey(code)) return label;
        string language = Culture.TwoLetterISOLanguageName;
        if (language == "en") {
            return Math.Abs(quantity) == 1M ? label : label + "s";
        }

        if (!UnitPluralForms.TryGetValue(code, out Dictionary<string, string[]>? byLanguage) ||
            !byLanguage.TryGetValue(language, out string[]? forms)) {
            return label;
        }

        return SelectUnitForm(language, Math.Abs(quantity), label, forms);
    }

    // Forms per language follow the CLDR cardinal categories that invoice quantities need:
    // de, es, it, nl, pt: { other }; fr: { other } with "one" below 2;
    // pl: { few, many, fraction }; cs and sk: { few, fraction, other }.
    private static string SelectUnitForm(string language, decimal quantity, string one, string[] forms) {
        bool integer = quantity == decimal.Truncate(quantity);
        switch (language) {
            case "fr":
                return quantity < 2M ? one : forms[0];
            case "pl": {
                if (!integer) return forms[2];
                if (quantity == 1M) return one;
                decimal mod10 = quantity % 10M, mod100 = quantity % 100M;
                return mod10 >= 2M && mod10 <= 4M && (mod100 < 12M || mod100 > 14M) ? forms[0] : forms[1];
            }
            case "cs":
            case "sk":
                if (!integer) return forms[1];
                if (quantity == 1M) return one;
                return quantity >= 2M && quantity <= 4M ? forms[0] : forms[2];
            default:
                return quantity == 1M ? one : forms[0];
        }
    }

    private static readonly Dictionary<string, Dictionary<string, string[]>> UnitPluralForms = new Dictionary<string, Dictionary<string, string[]>>(StringComparer.Ordinal) {
        ["C62"] = UnitForms(de: "Stück", fr: "pièces", es: "unidades", it: "pezzi", nl: "stuks", pt: "unidades",
            pl: new[] { "sztuki", "sztuk", "sztuki" }, cs: new[] { "kusy", "kusu", "kusů" }, sk: new[] { "kusy", "kusa", "kusov" }),
        ["HUR"] = UnitForms(de: "Stunden", fr: "heures", es: "horas", it: "ore", nl: "uur", pt: "horas",
            pl: new[] { "godziny", "godzin", "godziny" }, cs: new[] { "hodiny", "hodiny", "hodin" }, sk: new[] { "hodiny", "hodiny", "hodín" }),
        ["DAY"] = UnitForms(de: "Tage", fr: "jours", es: "días", it: "giorni", nl: "dagen", pt: "dias",
            pl: new[] { "dni", "dni", "dnia" }, cs: new[] { "dny", "dne", "dní" }, sk: new[] { "dni", "dňa", "dní" }),
        ["WEE"] = UnitForms(de: "Wochen", fr: "semaines", es: "semanas", it: "settimane", nl: "weken", pt: "semanas",
            pl: new[] { "tygodnie", "tygodni", "tygodnia" }, cs: new[] { "týdny", "týdne", "týdnů" }, sk: new[] { "týždne", "týždňa", "týždňov" }),
        ["MON"] = UnitForms(de: "Monate", fr: "mois", es: "meses", it: "mesi", nl: "maanden", pt: "meses",
            pl: new[] { "miesiące", "miesięcy", "miesiąca" }, cs: new[] { "měsíce", "měsíce", "měsíců" }, sk: new[] { "mesiace", "mesiaca", "mesiacov" }),
        ["KGM"] = UnitForms(de: "Kilogramm", fr: "kilogrammes", es: "kilogramos", it: "chilogrammi", nl: "kilogram", pt: "quilogramas",
            pl: new[] { "kilogramy", "kilogramów", "kilograma" }, cs: new[] { "kilogramy", "kilogramu", "kilogramů" }, sk: new[] { "kilogramy", "kilogramu", "kilogramov" }),
        ["MTR"] = UnitForms(de: "Meter", fr: "mètres", es: "metros", it: "metri", nl: "meter", pt: "metros",
            pl: new[] { "metry", "metrów", "metra" }, cs: new[] { "metry", "metru", "metrů" }, sk: new[] { "metre", "metra", "metrov" }),
        ["MTK"] = UnitForms(de: "Quadratmeter", fr: "mètres carrés", es: "metros cuadrados", it: "metri quadrati", nl: "vierkante meter", pt: "metros quadrados",
            pl: new[] { "metry kwadratowe", "metrów kwadratowych", "metra kwadratowego" },
            cs: new[] { "metry čtvereční", "metru čtverečního", "metrů čtverečních" },
            sk: new[] { "metre štvorcové", "metra štvorcového", "metrov štvorcových" }),
        ["LTR"] = UnitForms(de: "Liter", fr: "litres", es: "litros", it: "litri", nl: "liter", pt: "litros",
            pl: new[] { "litry", "litrów", "litra" }, cs: new[] { "litry", "litru", "litrů" }, sk: new[] { "litre", "litra", "litrov" })
    };

    private static Dictionary<string, string[]> UnitForms(string de, string fr, string es, string it, string nl, string pt, string[] pl, string[] cs, string[] sk) =>
        new Dictionary<string, string[]>(StringComparer.Ordinal) {
            ["de"] = new[] { de }, ["fr"] = new[] { fr }, ["es"] = new[] { es }, ["it"] = new[] { it },
            ["nl"] = new[] { nl }, ["pt"] = new[] { pt }, ["pl"] = pl, ["cs"] = cs, ["sk"] = sk
        };

    /// <summary>Returns the translated description of a supported UNCL 4461 payment code, or null when unknown.</summary>
    public string? GetPaymentDescription(string code) => CodeLabel(code, _paymentDescriptions, PaymentDescriptions);

    /// <summary>Returns the translated description of a UNCL 5305 VAT category code, or null when unknown.</summary>
    internal string? GetVatCategoryDescription(string code) => CodeLabel(code, null, VatCategoryDescriptions);

    /// <summary>Returns the translated description of a UNCL 1001 document type code, or null when unknown.</summary>
    internal string? GetDocumentTypeDescription(string code) => CodeLabel(code, null, DocumentTypeDescriptions);

    /// <summary>Creates an independent pack with caller-supplied code descriptions overriding built-in descriptions. Each dictionary accepts up to 512 entries, with codes up to 32 and descriptions up to 256 characters.</summary>
    public InvoicePdfLanguagePack WithCodeDescriptions(IReadOnlyDictionary<string, string>? units = null, IReadOnlyDictionary<string, string>? payments = null) =>
        new InvoicePdfLanguagePack(Culture, _labels, _builtInAdditionalLabels) {
            _unitDescriptions = units == null ? _unitDescriptions : CaptureDescriptions(units),
            _paymentDescriptions = payments == null ? _paymentDescriptions : CaptureDescriptions(payments)
        };

    private static ReadOnlyDictionary<string, string> CaptureDescriptions(IReadOnlyDictionary<string, string> source) {
        if (source.Count > 512) throw new ArgumentException("At most 512 code descriptions are supported.", nameof(source));
        var copy = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in source) {
            if (string.IsNullOrWhiteSpace(entry.Key) || entry.Key.Length > 32 || string.IsNullOrWhiteSpace(entry.Value) || entry.Value.Length > 256)
                throw new ArgumentException("Code descriptions require a nonempty code up to 32 and description up to 256 characters.", nameof(source));
            copy.Add(entry.Key, entry.Value.Trim());
        }
        return new ReadOnlyDictionary<string, string>(copy);
    }

    private string? CodeLabel(string code, IReadOnlyDictionary<string, string>? custom, IReadOnlyDictionary<string, string[]> builtIn) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(code);
#else
        if (code == null) throw new ArgumentNullException(nameof(code));
#endif
        if (custom != null && custom.TryGetValue(code, out string? value)) return value;
        if (!builtIn.TryGetValue(code, out string[]? labels)) return null;
        int index = Culture.TwoLetterISOLanguageName switch {
            "de" => 1, "pl" => 2, "fr" => 3, "es" => 4, "it" => 5, "nl" => 6, "pt" => 7, "cs" => 8, "sk" => 9, _ => 0
        };
        return labels[index];
    }

    // English, German, Polish, French, Spanish, Italian, Dutch, Portuguese, Czech, Slovak.
    private static readonly Dictionary<string, string[]> UnitDescriptions = new Dictionary<string, string[]>(StringComparer.Ordinal) {
        ["C62"] = new[] { "piece", "Stück", "sztuka", "pièce", "unidad", "pezzo", "stuk", "unidade", "kus", "kus" },
        ["HUR"] = new[] { "hour", "Stunde", "godzina", "heure", "hora", "ora", "uur", "hora", "hodina", "hodina" },
        ["DAY"] = new[] { "day", "Tag", "dzień", "jour", "día", "giorno", "dag", "dia", "den", "deň" },
        ["WEE"] = new[] { "week", "Woche", "tydzień", "semaine", "semana", "settimana", "week", "semana", "týden", "týždeň" },
        ["MON"] = new[] { "month", "Monat", "miesiąc", "mois", "mes", "mese", "maand", "mês", "měsíc", "mesiac" },
        ["KGM"] = new[] { "kilogram", "Kilogramm", "kilogram", "kilogramme", "kilogramo", "chilogrammo", "kilogram", "quilograma", "kilogram", "kilogram" },
        ["MTR"] = new[] { "metre", "Meter", "metr", "mètre", "metro", "metro", "meter", "metro", "metr", "meter" },
        ["MTK"] = new[] { "square metre", "Quadratmeter", "metr kwadratowy", "mètre carré", "metro cuadrado", "metro quadrato", "vierkante meter", "metro quadrado", "metr čtvereční", "meter štvorcový" },
        ["LTR"] = new[] { "litre", "Liter", "litr", "litre", "litro", "litro", "liter", "litro", "litr", "liter" }
    };

    private static readonly IReadOnlyDictionary<string, string[]> PaymentDescriptions = new Dictionary<string, string[]>(StringComparer.Ordinal) {
        ["10"] = new[] { "Cash", "Barzahlung", "Gotówka", "Espèces", "Efectivo", "Contanti", "Contant", "Numerário", "Hotovost", "Hotovosť" },
        ["30"] = new[] { "Credit transfer", "Überweisung", "Przelew", "Virement", "Transferencia", "Bonifico", "Overschrijving", "Transferência", "Bankovní převod", "Bankový prevod" },
        ["48"] = new[] { "Bank card", "Bankkarte", "Karta płatnicza", "Carte bancaire", "Tarjeta bancaria", "Carta bancaria", "Bankkaart", "Cartão bancário", "Platební karta", "Platobná karta" },
        ["49"] = new[] { "Direct debit", "Lastschrift", "Polecenie zapłaty", "Prélèvement", "Domiciliación", "Addebito diretto", "Automatische incasso", "Débito direto", "Inkaso", "Inkaso" },
        ["58"] = new[] { "SEPA credit transfer", "SEPA-Überweisung", "Przelew SEPA", "Virement SEPA", "Transferencia SEPA", "Bonifico SEPA", "SEPA-overschrijving", "Transferência SEPA", "Převod SEPA", "Prevod SEPA" },
        ["59"] = new[] { "SEPA direct debit", "SEPA-Lastschrift", "Polecenie zapłaty SEPA", "Prélèvement SEPA", "Domiciliación SEPA", "Addebito diretto SEPA", "SEPA-incasso", "Débito direto SEPA", "Inkaso SEPA", "Inkaso SEPA" }
    };

    private static readonly IReadOnlyDictionary<string, string[]> VatCategoryDescriptions = new Dictionary<string, string[]>(StringComparer.Ordinal) {
        ["S"] = new[] { "Standard rate", "Normalsatz", "Stawka podstawowa", "Taux normal", "Tipo general", "Aliquota ordinaria", "Standaardtarief", "Taxa normal", "Základní sazba", "Základná sadzba" },
        ["AA"] = new[] { "Reduced rate", "Ermäßigter Satz", "Stawka obniżona", "Taux réduit", "Tipo reducido", "Aliquota ridotta", "Verlaagd tarief", "Taxa reduzida", "Snížená sazba", "Znížená sadzba" },
        ["Z"] = new[] { "Zero rate", "Nullsatz", "Stawka zerowa", "Taux zéro", "Tipo cero", "Aliquota zero", "Nultarief", "Taxa zero", "Nulová sazba", "Nulová sadzba" },
        ["E"] = new[] { "Exempt", "Steuerbefreit", "Zwolnione", "Exonéré", "Exento", "Esente", "Vrijgesteld", "Isento", "Osvobozeno", "Oslobodené" },
        ["AE"] = new[] { "Reverse charge", "Umkehr der Steuerschuld", "Odwrotne obciążenie", "Autoliquidation", "Inversión del sujeto pasivo", "Inversione contabile", "Btw verlegd", "Autoliquidação", "Přenesená daňová povinnost", "Prenesenie daňovej povinnosti" },
        ["K"] = new[] { "Intra-community supply", "Innergemeinschaftliche Lieferung", "Wewnątrzwspólnotowa dostawa", "Livraison intracommunautaire", "Entrega intracomunitaria", "Cessione intracomunitaria", "Intracommunautaire levering", "Fornecimento intracomunitário", "Dodání do jiného členského státu", "Dodanie do iného členského štátu" },
        ["G"] = new[] { "Export outside the EU", "Ausfuhr außerhalb der EU", "Eksport poza UE", "Exportation hors UE", "Exportación fuera de la UE", "Esportazione fuori UE", "Uitvoer buiten de EU", "Exportação fora da UE", "Vývoz mimo EU", "Vývoz mimo EÚ" },
        ["O"] = new[] { "Not subject to VAT", "Nicht steuerbar", "Niepodlegające VAT", "Hors champ de la TVA", "No sujeto a IVA", "Fuori campo IVA", "Niet onderworpen aan btw", "Não sujeito a IVA", "Nepodléhá DPH", "Nepodlieha DPH" }
    };

    private static readonly IReadOnlyDictionary<string, string[]> DocumentTypeDescriptions = new Dictionary<string, string[]>(StringComparer.Ordinal) {
        ["380"] = new[] { "Commercial invoice", "Handelsrechnung", "Faktura handlowa", "Facture commerciale", "Factura comercial", "Fattura commerciale", "Handelsfactuur", "Fatura comercial", "Obchodní faktura", "Obchodná faktúra" },
        ["381"] = new[] { "Credit note", "Gutschrift", "Nota uznaniowa", "Avoir", "Nota de crédito", "Nota di credito", "Creditnota", "Nota de crédito", "Dobropis", "Dobropis" },
        ["384"] = new[] { "Corrected invoice", "Korrigierte Rechnung", "Faktura korygująca", "Facture corrigée", "Factura corregida", "Fattura corretta", "Gecorrigeerde factuur", "Fatura corrigida", "Opravená faktura", "Opravená faktúra" },
        ["386"] = new[] { "Prepayment invoice", "Vorauszahlungsrechnung", "Faktura zaliczkowa", "Facture d'acompte", "Factura de anticipo", "Fattura di acconto", "Voorschotfactuur", "Fatura de adiantamento", "Zálohová faktura", "Zálohová faktúra" },
        ["389"] = new[] { "Self-billed invoice", "Selbstfakturierte Rechnung", "Samofakturowanie", "Autofacturation", "Autofactura", "Autofattura", "Zelf gefactureerde factuur", "Autofaturação", "Samofakturace", "Samofakturácia" },
        ["326"] = new[] { "Partial invoice", "Teilrechnung", "Faktura częściowa", "Facture partielle", "Factura parcial", "Fattura parziale", "Deelfactuur", "Fatura parcial", "Dílčí faktura", "Čiastková faktúra" }
    };
}
