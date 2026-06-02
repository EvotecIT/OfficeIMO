namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a small vector icon drawn inside a table cell before the cell text.
/// </summary>
public sealed class PdfCellIcon {
    private PdfCellIconKind _kind;
    private double _size = 8D;
    private double _offsetX;
    private double _offsetY;

    /// <summary>Icon shape to draw.</summary>
    public PdfCellIconKind Kind {
        get => _kind;
        set {
            if (value < PdfCellIconKind.Circle || value > PdfCellIconKind.CheckBoxChecked) {
                throw new System.ArgumentOutOfRangeException(nameof(value), value, "PDF table cell icon kind is not supported.");
            }

            _kind = value;
        }
    }

    /// <summary>Icon fill color.</summary>
    public PdfColor Color { get; set; } = PdfColor.Black;

    /// <summary>Icon size in points.</summary>
    public double Size {
        get => _size;
        set {
            if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentException("PDF table cell icon size must be a positive finite value.", nameof(Size));
            }

            _size = value;
        }
    }

    /// <summary>Additional horizontal icon offset in points after table-cell alignment is applied.</summary>
    public double OffsetX {
        get => _offsetX;
        set {
            ValidateOffset(value, nameof(OffsetX));
            _offsetX = value;
        }
    }

    /// <summary>Additional vertical icon offset in points after table-cell alignment is applied.</summary>
    public double OffsetY {
        get => _offsetY;
        set {
            ValidateOffset(value, nameof(OffsetY));
            _offsetY = value;
        }
    }

    /// <summary>Creates a copy of this table cell icon.</summary>
    public PdfCellIcon Clone() => new PdfCellIcon {
        Kind = Kind,
        Color = Color,
        Size = Size,
        OffsetX = OffsetX,
        OffsetY = OffsetY
    };

    private static void ValidateOffset(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException("PDF table cell icon offsets must be finite values.", paramName);
        }
    }
}
