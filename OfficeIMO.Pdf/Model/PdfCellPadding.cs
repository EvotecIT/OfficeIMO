namespace OfficeIMO.Pdf;

/// <summary>
/// Describes optional per-side padding overrides for one table cell.
/// </summary>
public sealed class PdfCellPadding {
    private double? _left;
    private double? _right;
    private double? _top;
    private double? _bottom;

    /// <summary>Optional left padding in points. When null the table style value is used.</summary>
    public double? Left {
        get => _left;
        set {
            ValidateOptionalPadding(value, nameof(Left));
            _left = value;
        }
    }

    /// <summary>Optional right padding in points. When null the table style value is used.</summary>
    public double? Right {
        get => _right;
        set {
            ValidateOptionalPadding(value, nameof(Right));
            _right = value;
        }
    }

    /// <summary>Optional top padding in points. When null the table style value is used.</summary>
    public double? Top {
        get => _top;
        set {
            ValidateOptionalPadding(value, nameof(Top));
            _top = value;
        }
    }

    /// <summary>Optional bottom padding in points. When null the table style value is used.</summary>
    public double? Bottom {
        get => _bottom;
        set {
            ValidateOptionalPadding(value, nameof(Bottom));
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
