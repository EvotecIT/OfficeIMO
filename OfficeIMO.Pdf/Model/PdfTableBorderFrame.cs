namespace OfficeIMO.Pdf;

/// <summary>Retains space around a document table's cell grid and an optional independent perimeter.</summary>
internal sealed class PdfTableBorderFrame {
    internal double Spacing { get; set; }
    internal PdfCellBorder? Border { get; set; }
    internal double HorizontalInset { get; set; }
    internal PdfColor? Background { get; set; }
    /// <summary>Contains the outward paint of collapsed cell borders in ordinary table flow.</summary>
    internal bool ReserveCellBorderOutsets { get; set; }

    internal PdfTableBorderFrame Clone() => new() {
        Spacing = Spacing, Border = Border?.Clone(), HorizontalInset = HorizontalInset, Background = Background,
        ReserveCellBorderOutsets = ReserveCellBorderOutsets
    };
}
