namespace OfficeIMO.Pdf;

/// <summary>Describes one standard property column in a generated document portfolio.</summary>
public sealed class PdfPortfolioField {
    private string _displayName;
    private int _order;

    /// <summary>Creates a portfolio field for a standard embedded-file property.</summary>
    public PdfPortfolioField(PdfPortfolioFieldKind kind, string displayName, int order = 0) {
        ValidateKind(kind, nameof(kind));
        Guard.NotNullOrWhiteSpace(displayName, nameof(displayName));
        Guard.NonNegative(order, nameof(order));
        Kind = kind;
        _displayName = displayName;
        _order = order;
    }

    /// <summary>Standard embedded-file property represented by this field.</summary>
    public PdfPortfolioFieldKind Kind { get; }

    /// <summary>Localized field label displayed by a compatible viewer.</summary>
    public string DisplayName {
        get => _displayName;
        set {
            Guard.NotNullOrWhiteSpace(value, nameof(DisplayName));
            _displayName = value;
        }
    }

    /// <summary>Zero-based display order.</summary>
    public int Order {
        get => _order;
        set {
            Guard.NonNegative(value, nameof(Order));
            _order = value;
        }
    }

    /// <summary>Whether compatible viewers should show the field.</summary>
    public bool Visible { get; set; } = true;

    /// <summary>Whether compatible viewers may edit the field.</summary>
    public bool Editable { get; set; }

    internal string Key => Kind switch {
        PdfPortfolioFieldKind.FileName => "FileName",
        PdfPortfolioFieldKind.Description => "Description",
        PdfPortfolioFieldKind.CreationDate => "CreationDate",
        PdfPortfolioFieldKind.ModificationDate => "ModificationDate",
        PdfPortfolioFieldKind.Size => "Size",
        _ => throw new ArgumentOutOfRangeException(nameof(Kind), Kind, "Unsupported PDF portfolio field kind.")
    };

    internal PdfPortfolioField Clone() => new PdfPortfolioField(Kind, DisplayName, Order) {
        Visible = Visible,
        Editable = Editable
    };

    internal static void ValidateKind(PdfPortfolioFieldKind kind, string paramName) {
        if (kind < PdfPortfolioFieldKind.FileName || kind > PdfPortfolioFieldKind.Size) {
            throw new ArgumentOutOfRangeException(paramName, kind, "Unsupported PDF portfolio field kind.");
        }
    }
}
