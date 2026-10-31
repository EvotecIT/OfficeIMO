using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

/// <summary>Defines the colors and geometry used by the optional modern invoice presentation.</summary>
public sealed class InvoicePdfTheme {
    private double _cornerRadius = 7D;

    /// <summary>Primary accent used for headings, table headers, and highlighted totals.</summary>
    public global::OfficeIMO.OfficeColor Accent { get; set; } = global::OfficeIMO.OfficeColor.FromRgb(63, 92, 255);

    /// <summary>Primary text color.</summary>
    public global::OfficeIMO.OfficeColor Text { get; set; } = global::OfficeIMO.OfficeColor.FromRgb(1, 21, 52);

    /// <summary>Secondary text color.</summary>
    public global::OfficeIMO.OfficeColor MutedText { get; set; } = global::OfficeIMO.OfficeColor.FromRgb(82, 99, 122);

    /// <summary>Soft background used for cards and alternating rows.</summary>
    public global::OfficeIMO.OfficeColor Surface { get; set; } = global::OfficeIMO.OfficeColor.FromRgb(245, 248, 252);

    /// <summary>Border and separator color.</summary>
    public global::OfficeIMO.OfficeColor Border { get; set; } = global::OfficeIMO.OfficeColor.FromRgb(214, 223, 235);

    /// <summary>Corner radius in points for cards and tables.</summary>
    public double CornerRadius {
        get => _cornerRadius;
        set {
            if (value < 0D || double.IsNaN(value) || double.IsInfinity(value))
                throw new ArgumentOutOfRangeException(nameof(value), "The invoice corner radius must be a non-negative finite value.");
            _cornerRadius = value;
        }
    }

    /// <summary>Creates a modern theme from one brand accent while retaining accessible neutral colors.</summary>
    public static InvoicePdfTheme Modern(global::OfficeIMO.OfficeColor accent) => new InvoicePdfTheme { Accent = accent };

    internal InvoicePdfTheme Snapshot() => new InvoicePdfTheme {
        Accent = Accent,
        Text = Text,
        MutedText = MutedText,
        Surface = Surface,
        Border = Border,
        CornerRadius = CornerRadius
    };
}
