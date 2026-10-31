namespace OfficeIMO.Pdf;

/// <summary>
/// Describes optional per-side padding overrides for one table cell.
/// </summary>
public sealed class PdfCellPadding {
    private global::OfficeIMO.OfficeLength? _left;
    private global::OfficeIMO.OfficeLength? _right;
    private global::OfficeIMO.OfficeLength? _top;
    private global::OfficeIMO.OfficeLength? _bottom;

    /// <summary>Optional left padding as a physical length. When null the table style value is used.</summary>
    public global::OfficeIMO.OfficeLength? Left {
        get => _left;
        set {
            ValidateOptionalPadding(value?.ToPoints(), nameof(Left));
            _left = value;
        }
    }

    /// <summary>Optional right padding as a physical length. When null the table style value is used.</summary>
    public global::OfficeIMO.OfficeLength? Right {
        get => _right;
        set {
            ValidateOptionalPadding(value?.ToPoints(), nameof(Right));
            _right = value;
        }
    }

    /// <summary>Optional top padding as a physical length. When null the table style value is used.</summary>
    public global::OfficeIMO.OfficeLength? Top {
        get => _top;
        set {
            ValidateOptionalPadding(value?.ToPoints(), nameof(Top));
            _top = value;
        }
    }

    /// <summary>Optional bottom padding as a physical length. When null the table style value is used.</summary>
    public global::OfficeIMO.OfficeLength? Bottom {
        get => _bottom;
        set {
            ValidateOptionalPadding(value?.ToPoints(), nameof(Bottom));
            _bottom = value;
        }
    }

    /// <summary>Creates a copy of this table cell padding override.</summary>
    public PdfCellPadding Clone() => new PdfCellPadding {
        Left = Left,
        Right = Right,
        Top = Top,
        Bottom = Bottom
    };

    private static void ValidateOptionalPadding(double? value, string paramName) {
        if (value.HasValue && (value.Value < 0 || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
            throw new System.ArgumentException("Table cell padding values must be non-negative finite values.", paramName);
        }
    }
}
