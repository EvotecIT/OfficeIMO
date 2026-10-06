namespace OfficeIMO.Pdf;

/// <summary>Retains a document table's independent perimeter and space around its cell grid.</summary>
internal sealed class PdfTableBorderFrame {
    internal double Spacing { get; set; }
    internal PdfCellBorder? Border { get; set; }
    internal double HorizontalInset { get; set; }
    internal PdfColor? Background { get; set; }

    internal PdfTableBorderFrame Clone() => new() {
        Spacing = Spacing, Border = Border?.Clone(), HorizontalInset = HorizontalInset, Background = Background
    };
}
