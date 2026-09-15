using System.Globalization;

namespace OfficeIMO.Invoicing.Pdf;

/// <summary>Controls generated labels and locale-sensitive formatting in the visible invoice PDF.</summary>
public sealed class InvoicePdfLayoutOptions {
    private readonly List<InvoicePdfLanguagePack> _languages = new List<InvoicePdfLanguagePack>();
    private string _labelSeparator = " / ";
    private string _dateFormat = "yyyy-MM-dd";
    private CultureInfo _formattingCulture = CultureInfo.InvariantCulture;

    /// <summary>Creates the compatibility layout: English labels with invariant numbers and ISO dates.</summary>
    public InvoicePdfLayoutOptions() => _languages.Add(InvoicePdfLanguagePack.ForCulture("en-US"));

    /// <summary>Language packs displayed in order. Multiple packs produce multilingual labels.</summary>
    public IList<InvoicePdfLanguagePack> Languages => _languages;

    /// <summary>Text placed between translations of the same generated label.</summary>
    public string LabelSeparator {
        get => _labelSeparator;
        set => _labelSeparator = string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException("A multilingual label separator is required.", nameof(value))
            : value;
    }

    /// <summary>Culture used for visible decimal and date formatting.</summary>
    public CultureInfo FormattingCulture {
        get => _formattingCulture;
        set => _formattingCulture = CultureInfo.ReadOnly((CultureInfo)(value ?? throw new ArgumentNullException(nameof(value))).Clone());
    }

    /// <summary>.NET date format used for visible dates. The default compatibility layout uses <c>yyyy-MM-dd</c>.</summary>
    public string DateFormat {
        get => _dateFormat;
        set => _dateFormat = string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException("A date format is required.", nameof(value))
            : value;
    }

    /// <summary>
    /// Creates a localized or multilingual layout from built-in language packs. The first culture formats numbers and dates.
    /// </summary>
    public static InvoicePdfLayoutOptions ForCultures(params string[] cultureNames) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(cultureNames);
#else
        if (cultureNames == null) throw new ArgumentNullException(nameof(cultureNames));
#endif
        if (cultureNames.Length == 0) throw new ArgumentException("At least one culture is required.", nameof(cultureNames));
        var result = new InvoicePdfLayoutOptions();
        result._languages.Clear();
        foreach (string cultureName in cultureNames) result._languages.Add(InvoicePdfLanguagePack.ForCulture(cultureName));
        result.FormattingCulture = result._languages[0].Culture;
        result.DateFormat = "d";
        return result;
    }

    internal InvoicePdfLayoutOptions Snapshot() {
        if (_languages.Count == 0) throw new InvalidOperationException("At least one invoice PDF language pack is required.");
        if (_languages.Any(static language => language == null)) throw new InvalidOperationException("Invoice PDF language packs cannot contain null.");
        var result = new InvoicePdfLayoutOptions {
            LabelSeparator = LabelSeparator,
            FormattingCulture = FormattingCulture,
            DateFormat = DateFormat
        };
        result._languages.Clear();
        result._languages.AddRange(_languages);
        return result;
    }
}
