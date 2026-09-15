using System.Collections.ObjectModel;
using System.Globalization;

namespace OfficeIMO.Invoicing.Pdf;

/// <summary>Defines the culture and generated labels for one invoice presentation language.</summary>
public sealed partial class InvoicePdfLanguagePack {
    private readonly IReadOnlyDictionary<InvoicePdfText, string> _labels;

    private InvoicePdfLanguagePack(CultureInfo culture, IReadOnlyDictionary<InvoicePdfText, string> labels) {
        Culture = CultureInfo.ReadOnly((CultureInfo)culture.Clone());
        _labels = labels;
    }

    /// <summary>Culture used by this language pack.</summary>
    public CultureInfo Culture { get; }

    /// <summary>Returns the localized label for <paramref name="text"/>.</summary>
    public string this[InvoicePdfText text] => _labels.TryGetValue(text, out string? value) ? value : EnglishLabels[text];

    /// <summary>
    /// Creates a pack for a supported built-in language. English, German, Polish, and French are supported;
    /// regional culture names use the matching language translation.
    /// </summary>
    public static InvoicePdfLanguagePack ForCulture(string cultureName) {
        if (string.IsNullOrWhiteSpace(cultureName)) throw new ArgumentException("A culture name is required.", nameof(cultureName));
        CultureInfo culture = CultureInfo.GetCultureInfo(cultureName.Trim());
        IReadOnlyDictionary<InvoicePdfText, string> labels = culture.TwoLetterISOLanguageName switch {
            "en" => EnglishLabels,
            "de" => GermanLabels,
            "pl" => PolishLabels,
            "fr" => FrenchLabels,
            _ => throw new NotSupportedException("Built-in invoice PDF labels support English, German, Polish, and French. Use Create for another language.")
        };
        return new InvoicePdfLanguagePack(culture, labels);
    }

    /// <summary>
    /// Creates a custom pack. Missing entries deliberately fall back to English so callers can override only the labels they own.
    /// </summary>
    public static InvoicePdfLanguagePack Create(string cultureName, IReadOnlyDictionary<InvoicePdfText, string> labels) {
        if (string.IsNullOrWhiteSpace(cultureName)) throw new ArgumentException("A culture name is required.", nameof(cultureName));
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(labels);
#else
        if (labels == null) throw new ArgumentNullException(nameof(labels));
#endif
        var copy = new Dictionary<InvoicePdfText, string>();
        foreach (KeyValuePair<InvoicePdfText, string> entry in labels) {
            if (entry.Key < InvoicePdfText.Invoice || entry.Key > InvoicePdfText.SupportingDocuments) throw new ArgumentOutOfRangeException(nameof(labels), "A label key is undefined.");
            if (string.IsNullOrWhiteSpace(entry.Value)) throw new ArgumentException("Invoice PDF labels cannot be empty.", nameof(labels));
            copy[entry.Key] = entry.Value.Trim();
        }
        return new InvoicePdfLanguagePack(CultureInfo.GetCultureInfo(cultureName.Trim()),
            new ReadOnlyDictionary<InvoicePdfText, string>(copy));
    }
}
