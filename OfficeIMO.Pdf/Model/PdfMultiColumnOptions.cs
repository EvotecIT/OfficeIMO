namespace OfficeIMO.Pdf;

/// <summary>Controls automatic block flow across equal-width page columns.</summary>
public sealed class PdfMultiColumnOptions {
    private int _columnCount = 2;
    private double _gap = PdfRowStyle.DefaultGap;
    private double _separatorWidth;

    /// <summary>Number of equal-width columns.</summary>
    public int ColumnCount {
        get => _columnCount;
        set {
            if (value < 2 || value > 12) throw new ArgumentOutOfRangeException(nameof(ColumnCount), value, "PDF multi-column layouts require between 2 and 12 columns.");
            _columnCount = value;
        }
    }

    /// <summary>Horizontal gutter between columns in points.</summary>
    public double Gap {
        get => _gap;
        set {
            if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(Gap), value, "PDF multi-column gap must be non-negative and finite.");
            _gap = value;
        }
    }

    /// <summary>Whether the final page should distribute blocks toward similar column heights.</summary>
    public bool BalanceLastPage { get; set; } = true;
    /// <summary>Whether long paragraphs may be split at already wrapped line boundaries to balance the final page.</summary>
    public bool BalanceParagraphLines { get; set; } = true;
    /// <summary>Optional separator color between columns.</summary>
    public PdfColor? SeparatorColor { get; set; }
    /// <summary>Separator width in points.</summary>
    public double SeparatorWidth {
        get => _separatorWidth;
        set {
            if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(SeparatorWidth), value, "PDF multi-column separator width must be non-negative and finite.");
            _separatorWidth = value;
        }
    }

    internal PdfMultiColumnOptions Clone() => new PdfMultiColumnOptions {
        ColumnCount = ColumnCount,
        Gap = Gap,
        BalanceLastPage = BalanceLastPage,
        BalanceParagraphLines = BalanceParagraphLines,
        SeparatorColor = SeparatorColor,
        SeparatorWidth = SeparatorWidth
    };
}
