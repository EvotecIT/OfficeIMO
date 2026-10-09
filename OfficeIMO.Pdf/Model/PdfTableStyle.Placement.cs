namespace OfficeIMO.Pdf;

public partial class PdfTableStyle {
    private double _leftIndent;
    private double _horizontalOffset;

    /// <summary>Left indentation before table placement, in points. Negative values extend the table into the leading margin.</summary>
    public double LeftIndent {
        get => _leftIndent;
        set {
            ValidateFiniteValue(value, nameof(LeftIndent), "Table left indent must be a finite value.");
            _leftIndent = value;
        }
    }

    /// <summary>
    /// Additional horizontal translation after table alignment or floating placement, in points.
    /// Positive values move the table right. This does not change its available width or vertical flow.
    /// </summary>
    public double HorizontalOffset {
        get => _horizontalOffset;
        set {
            ValidateFiniteValue(value, nameof(HorizontalOffset), "Table horizontal offset must be a finite value.");
            _horizontalOffset = value;
        }
    }
}
