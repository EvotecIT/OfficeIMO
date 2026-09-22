using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

/// <summary>Controls generated labels and locale-sensitive formatting in the visible invoice PDF.</summary>
public sealed class InvoicePdfLayoutOptions {
    private readonly List<InvoicePdfLanguagePack> _languages = new List<InvoicePdfLanguagePack>();
    private readonly List<InvoicePdfApproval> _approvals = new List<InvoicePdfApproval>();
    private string _labelSeparator = " / ";
    private string _dateFormat = "yyyy-MM-dd";
    private CultureInfo _formattingCulture = CultureInfo.InvariantCulture;
    private byte[]? _logoBytes;
    private double _logoWidth = 150D;
    private double _logoHeight = 30D;
    private OfficeImageFit _logoFit = OfficeImageFit.Contain;
    private int _maxInvoiceXmlBytes = 256 * 1024;
    private int _maxLineTextCharacters = 4_096;
    private int _maxInvoiceLines = 1_000;
    private int _maxGeneratedPages = 100;
    private int _maxOutputBytes = 32 * 1024 * 1024;

    /// <summary>Creates the compatibility layout: English labels with invariant numbers and ISO dates.</summary>
    public InvoicePdfLayoutOptions() => _languages.Add(InvoicePdfLanguagePack.ForCulture("en-US"));

    /// <summary>Maximum source invoice XML bytes accepted for PDF presentation.</summary>
    public int MaxInvoiceXmlBytes { get => _maxInvoiceXmlBytes; set => _maxInvoiceXmlBytes = Positive(value, nameof(value)); }
    /// <summary>Maximum visible text characters in a single invoice line item.</summary>
    public int MaxLineTextCharacters { get => _maxLineTextCharacters; set => _maxLineTextCharacters = Positive(value, nameof(value)); }
    /// <summary>Maximum invoice line items accepted for PDF presentation.</summary>
    public int MaxInvoiceLines { get => _maxInvoiceLines; set => _maxInvoiceLines = Positive(value, nameof(value)); }
    /// <summary>Maximum pages generated during invoice PDF layout.</summary>
    public int MaxGeneratedPages { get => _maxGeneratedPages; set => _maxGeneratedPages = Positive(value, nameof(value)); }
    /// <summary>Maximum serialized PDF bytes returned by invoice rendering.</summary>
    public int MaxOutputBytes { get => _maxOutputBytes; set => _maxOutputBytes = Positive(value, nameof(value)); }

    /// <summary>Language packs displayed in order. Multiple packs produce multilingual labels.</summary>
    public IList<InvoicePdfLanguagePack> Languages => _languages;

    /// <summary>
    /// Optional visual theme. Leave null to retain the compact compatibility layout.
    /// </summary>
    public InvoicePdfTheme? Theme { get; set; }

    /// <summary>Optional raster logo bytes shown in the modern header. Values are copied defensively.</summary>
    public byte[]? LogoBytes {
        get => _logoBytes == null ? null : (byte[])_logoBytes.Clone();
        set {
            if (value != null && value.Length == 0) throw new ArgumentException("Logo bytes cannot be empty.", nameof(value));
            _logoBytes = value == null ? null : (byte[])value.Clone();
        }
    }

    /// <summary>Logo width in points. Defaults to 150.</summary>
    public double LogoWidth {
        get => _logoWidth;
        set => _logoWidth = PositiveFinite(value, nameof(value));
    }

    /// <summary>Logo height in points. Defaults to 30.</summary>
    public double LogoHeight {
        get => _logoHeight;
        set => _logoHeight = PositiveFinite(value, nameof(value));
    }

    /// <summary>How the logo fits its width and height box. Defaults to <see cref="OfficeImageFit.Contain"/> to preserve aspect ratio.</summary>
    public OfficeImageFit LogoFit {
        get => _logoFit;
        set {
            if (value != OfficeImageFit.Stretch && value != OfficeImageFit.Contain && value != OfficeImageFit.Cover)
                throw new ArgumentOutOfRangeException(nameof(value));
            _logoFit = value;
        }
    }

    /// <summary>Alternate text for the logo when tagged PDF output is enabled.</summary>
    public string? LogoAlternativeText { get; set; }

    /// <summary>
    /// Visible prepared/approved blocks placed after payment details. They do not create cryptographic PDF signatures.
    /// </summary>
    public IList<InvoicePdfApproval> Approvals => _approvals;

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
            DateFormat = DateFormat,
            Theme = Theme?.Snapshot(),
            LogoBytes = _logoBytes,
            LogoWidth = LogoWidth,
            LogoHeight = LogoHeight,
            LogoFit = LogoFit,
            LogoAlternativeText = LogoAlternativeText,
            MaxInvoiceXmlBytes = MaxInvoiceXmlBytes,
            MaxLineTextCharacters = MaxLineTextCharacters,
            MaxInvoiceLines = MaxInvoiceLines,
            MaxGeneratedPages = MaxGeneratedPages,
            MaxOutputBytes = MaxOutputBytes
        };
        result._languages.Clear();
        result._languages.AddRange(_languages);
        result._approvals.AddRange(_approvals.Select(static approval =>
            approval?.Snapshot() ?? throw new InvalidOperationException("Invoice approval blocks cannot contain null.")));
        return result;
    }

    internal byte[]? LogoBytesSnapshot => _logoBytes;

    private static double PositiveFinite(double value, string paramName) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value))
            throw new ArgumentOutOfRangeException(paramName, "The value must be a positive finite number.");
        return value;
    }

    private static int Positive(int value, string paramName) => value > 0
        ? value : throw new ArgumentOutOfRangeException(paramName);
}
