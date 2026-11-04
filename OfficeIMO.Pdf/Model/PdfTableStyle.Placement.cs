namespace OfficeIMO.Pdf;

public partial class PdfTableStyle {
    private global::OfficeIMO.OfficeLength _leftIndent;
    private global::OfficeIMO.OfficeLength _horizontalOffset;

    /// <summary>Document adapters can preserve an authored fixed grid beyond its containing text frame.</summary>
    internal bool AllowWidthBeyondFrame { get; set; }

    /// <summary>Left indentation before table placement, as a physical length. Negative values extend the table into the leading margin.</summary>
    public global::OfficeIMO.OfficeLength LeftIndent {
        get => _leftIndent;
        set {
            ValidateFiniteValue(value.ToPoints(), nameof(LeftIndent), "Table left indent must be a finite value.");
            _leftIndent = value;
        }
    }

    /// <summary>
    /// Additional horizontal translation after table alignment or floating placement, as a physical length.
    /// Positive values move the table right. This does not change its available width or vertical flow.
    /// </summary>
    public global::OfficeIMO.OfficeLength HorizontalOffset {
        get => _horizontalOffset;
        set {
            ValidateFiniteValue(value.ToPoints(), nameof(HorizontalOffset), "Table horizontal offset must be a finite value.");
            _horizontalOffset = value;
        }
    }
}
