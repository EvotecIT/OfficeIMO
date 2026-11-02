namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a small vector icon drawn inside a table cell before the cell text.
/// </summary>
public sealed class PdfCellIcon {
    private PdfCellIconKind _kind;
    private global::OfficeIMO.OfficeLength _size = global::OfficeIMO.OfficeLength.Points(8D);
    private global::OfficeIMO.OfficeLength _offsetX;
    private global::OfficeIMO.OfficeLength _offsetY;
    private PdfColumnAlign? _horizontalAlignment;

    /// <summary>Optional icon alignment independent of cell text. Null uses the cell's alignment.</summary>
    public PdfColumnAlign? HorizontalAlignment {
        get => _horizontalAlignment;
        set {
            if (value.HasValue && (value.Value < PdfColumnAlign.Left || value.Value > PdfColumnAlign.Right))
                throw new System.ArgumentOutOfRangeException(nameof(value), "Icon alignment must be a supported table column alignment.");
            _horizontalAlignment = value;
        }
    }

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
    public global::OfficeIMO.OfficeColor Color { get; set; } = global::OfficeIMO.OfficeColor.Black;

    /// <summary>Icon size as a physical length.</summary>
    public global::OfficeIMO.OfficeLength Size {
        get => _size;
        set {
            if (value.ToPoints() <= 0D || double.IsNaN(value.ToPoints()) || double.IsInfinity(value.ToPoints())) {
                throw new System.ArgumentException("PDF table cell icon size must be a positive finite value.", nameof(Size));
            }

            _size = value;
        }
    }

    /// <summary>Additional horizontal icon offset as a physical length after table-cell alignment is applied.</summary>
    public global::OfficeIMO.OfficeLength OffsetX {
        get => _offsetX;
        set {
            ValidateOffset(value.ToPoints(), nameof(OffsetX));
            _offsetX = value;
        }
    }

    /// <summary>Additional vertical icon offset as a physical length after table-cell alignment is applied.</summary>
    public global::OfficeIMO.OfficeLength OffsetY {
        get => _offsetY;
        set {
            ValidateOffset(value.ToPoints(), nameof(OffsetY));
            _offsetY = value;
        }
    }

    /// <summary>Creates a copy of this table cell icon.</summary>
    public PdfCellIcon Clone() => new PdfCellIcon {
        Kind = Kind,
        Color = Color,
        Size = Size,
        OffsetX = OffsetX,
        OffsetY = OffsetY,
        HorizontalAlignment = HorizontalAlignment
    };

    private static void ValidateOffset(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException("PDF table cell icon offsets must be finite values.", paramName);
        }
    }
}
