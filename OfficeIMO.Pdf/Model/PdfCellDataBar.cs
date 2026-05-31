namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a proportional visual bar drawn inside a table cell, behind the cell text.
/// </summary>
public sealed class PdfCellDataBar {
    private double _ratio;

    /// <summary>Fill color used for the data bar.</summary>
    public PdfColor Color { get; set; } = PdfColor.LightGray;

    /// <summary>Filled width as a 0..1 fraction of the cell content width.</summary>
    public double Ratio {
        get => _ratio;
        set {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0 || value > 1) {
                throw new ArgumentOutOfRangeException(nameof(Ratio), "PDF table data bar ratio must be a finite value between 0 and 1.");
            }

            _ratio = value;
        }
    }

    /// <summary>Creates a deep copy of this data bar.</summary>
    public PdfCellDataBar Clone() {
        return new PdfCellDataBar {
            Color = Color,
            Ratio = Ratio
        };
    }
}
