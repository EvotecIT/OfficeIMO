namespace OfficeIMO.Pdf;

/// <summary>
/// Describes an override border for a single table cell.
/// </summary>
public sealed class PdfCellBorder {
    private double _width = 0.5;
    private PdfCellBorderSide? _topBorder;
    private PdfCellBorderSide? _rightBorder;
    private PdfCellBorderSide? _bottomBorder;
    private PdfCellBorderSide? _leftBorder;
    private PdfCellBorderSide? _diagonalUpBorder;
    private PdfCellBorderSide? _diagonalDownBorder;

    /// <summary>Border color. Set to null or use a zero width to suppress the border.</summary>
    public PdfColor? Color { get; set; } = new PdfColor(0.8, 0.8, 0.8);

    /// <summary>Border stroke width in points.</summary>
    public double Width {
        get => _width;
        set {
            if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentException("Table cell border widths must be non-negative finite values.", nameof(Width));
            }

            _width = value;
        }
    }

    /// <summary>Border stroke dash style used by sides without an explicit side override.</summary>
    public OfficeIMO.Drawing.OfficeStrokeDashStyle DashStyle { get; set; }

    /// <summary>Border line style used by sides without an explicit side override.</summary>
    public PdfCellBorderLineStyle LineStyle { get; set; }

    /// <summary>Optional top border override. When set, it overrides the shared color and width for this side.</summary>
    public PdfCellBorderSide? TopBorder {
        get => _topBorder?.Clone();
        set => _topBorder = value?.Clone();
    }

    /// <summary>Optional right border override. When set, it overrides the shared color and width for this side.</summary>
    public PdfCellBorderSide? RightBorder {
        get => _rightBorder?.Clone();
        set => _rightBorder = value?.Clone();
    }

    /// <summary>Optional bottom border override. When set, it overrides the shared color and width for this side.</summary>
    public PdfCellBorderSide? BottomBorder {
        get => _bottomBorder?.Clone();
        set => _bottomBorder = value?.Clone();
    }

    /// <summary>Optional left border override. When set, it overrides the shared color and width for this side.</summary>
    public PdfCellBorderSide? LeftBorder {
        get => _leftBorder?.Clone();
        set => _leftBorder = value?.Clone();
    }

    /// <summary>Optional diagonal-up border override. The diagonal-up line runs from the bottom-left corner to the top-right corner.</summary>
    public PdfCellBorderSide? DiagonalUpBorder {
        get => _diagonalUpBorder?.Clone();
        set => _diagonalUpBorder = value?.Clone();
    }

    /// <summary>Optional diagonal-down border override. The diagonal-down line runs from the top-left corner to the bottom-right corner.</summary>
    public PdfCellBorderSide? DiagonalDownBorder {
        get => _diagonalDownBorder?.Clone();
        set => _diagonalDownBorder = value?.Clone();
    }

    internal PdfCellBorderSide? TopBorderSnapshot => _topBorder;
    internal PdfCellBorderSide? RightBorderSnapshot => _rightBorder;
    internal PdfCellBorderSide? BottomBorderSnapshot => _bottomBorder;
    internal PdfCellBorderSide? LeftBorderSnapshot => _leftBorder;
    internal PdfCellBorderSide? DiagonalUpBorderSnapshot => _diagonalUpBorder;
    internal PdfCellBorderSide? DiagonalDownBorderSnapshot => _diagonalDownBorder;

    /// <summary>Whether to draw the top side of the cell border.</summary>
    public bool Top { get; set; } = true;

    /// <summary>Whether to draw the right side of the cell border.</summary>
    public bool Right { get; set; } = true;

    /// <summary>Whether to draw the bottom side of the cell border.</summary>
    public bool Bottom { get; set; } = true;

    /// <summary>Whether to draw the left side of the cell border.</summary>
    public bool Left { get; set; } = true;

    /// <summary>Whether to draw the diagonal-up line from bottom-left to top-right.</summary>
    public bool DiagonalUp { get; set; }

    /// <summary>Whether to draw the diagonal-down line from top-left to bottom-right.</summary>
    public bool DiagonalDown { get; set; }

    /// <summary>Creates a deep copy of this border style.</summary>
    public PdfCellBorder Clone() => new PdfCellBorder {
        Color = Color,
        Width = Width,
        TopBorder = _topBorder,
        RightBorder = _rightBorder,
        BottomBorder = _bottomBorder,
        LeftBorder = _leftBorder,
        DiagonalUpBorder = _diagonalUpBorder,
        DiagonalDownBorder = _diagonalDownBorder,
        DashStyle = DashStyle,
        LineStyle = LineStyle,
        Top = Top,
        Right = Right,
        Bottom = Bottom,
        Left = Left,
        DiagonalUp = DiagonalUp,
        DiagonalDown = DiagonalDown
    };
}
