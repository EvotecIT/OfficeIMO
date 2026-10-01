using System.Collections.ObjectModel;

namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class InvoicePdfLanguagePack {
    private IReadOnlyDictionary<string, string>? _unitDescriptions;
    private IReadOnlyDictionary<string, string>? _paymentDescriptions;

    /// <summary>Returns the translated description of a supported UN/ECE unit code, or null when unknown.</summary>
    public string? GetUnitDescription(string code) => CodeLabel(code, _unitDescriptions, UnitDescriptions);

    /// <summary>Returns the translated description of a supported UNCL 4461 payment code, or null when unknown.</summary>
    public string? GetPaymentDescription(string code) => CodeLabel(code, _paymentDescriptions, PaymentDescriptions);

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
    private static readonly IReadOnlyDictionary<string, string[]> UnitDescriptions = new Dictionary<string, string[]>(StringComparer.Ordinal) {
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
}
