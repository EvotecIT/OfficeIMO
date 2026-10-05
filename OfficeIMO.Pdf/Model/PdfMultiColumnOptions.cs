namespace OfficeIMO.Pdf;

/// <summary>Controls sequential block flow across page columns.</summary>
public sealed class PdfMultiColumnOptions {
    private int _columnCount = 2;
    private double _gap = PdfRowStyle.DefaultGap;
    private double _separatorWidth;
    private double _finalColumnSpacingAfter;
    private IReadOnlyList<PdfFlowColumn> _columnDefinitions = Array.Empty<PdfFlowColumn>();

    /// <summary>Number of columns. Explicit definitions set this count and must be cleared before changing it.</summary>
    public int ColumnCount {
        get => _columnCount;
        set {
            if (value < 2 || value > 12) throw new ArgumentOutOfRangeException(nameof(ColumnCount), value, "PDF multi-column layouts require between 2 and 12 columns.");
            if (_columnDefinitions.Count > 0 && value != _columnDefinitions.Count)
                throw new InvalidOperationException("Replace or clear ColumnDefinitions before changing the column count.");
            _columnCount = value;
        }
    }

    /// <summary>Optional explicit widths and individual gutters. The setter snapshots the sequence; an empty sequence restores equal widths.</summary>
    public IReadOnlyList<PdfFlowColumn> ColumnDefinitions {
        get => _columnDefinitions;
        set {
            Guard.NotNull(value, nameof(value));
            PdfFlowColumn[] definitions = value.ToArray();
            if (definitions.Length > 0 && (definitions.Length < 2 || definitions.Length > 12))
                throw new ArgumentOutOfRangeException(nameof(value), "PDF multi-column layouts require between 2 and 12 columns.");
            if (definitions.Any(column => column == null))
                throw new ArgumentException("Column definitions cannot contain null entries.", nameof(value));
            _columnDefinitions = Array.AsReadOnly(definitions);
            if (definitions.Length > 0) _columnCount = definitions.Length;
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
    /// <summary>Allows kept paragraphs longer than a balanced column to split on the same physical page.
    /// Paragraphs that fit a balanced column and full-height sequential frames retain their keep-together rules. Defaults to false.</summary>
    public bool BalanceKeptParagraphLines { get; set; }
    /// <summary>Preserves the terminal keep-with-next group when balancing shortened column frames. Defaults to true.
    /// Full-height and partial physical-page frames retain their ordinary keep rules.</summary>
    public bool HonorKeepWithNextWhenBalancing { get; set; } = true;
    /// <summary>Whether splittable table rows may be balanced at legal cell paragraph boundaries. Defaults to whole-row balancing.</summary>
    public bool BalanceTableRowLines { get; set; }
    /// <summary>Space below the final occupied column after balancing, in points. Limited to the remaining physical page space.</summary>
    public double FinalColumnSpacingAfter {
        get => _finalColumnSpacingAfter;
        set {
            if (value < 0 || double.IsNaN(value) || double.IsInfinity(value))
                throw new ArgumentOutOfRangeException(nameof(FinalColumnSpacingAfter), value, "Final column spacing must be non-negative and finite.");
            _finalColumnSpacingAfter = value;
        }
    }
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
        ColumnDefinitions = ColumnDefinitions,
        Gap = Gap,
        BalanceLastPage = BalanceLastPage,
        BalanceParagraphLines = BalanceParagraphLines,
        BalanceKeptParagraphLines = BalanceKeptParagraphLines,
        HonorKeepWithNextWhenBalancing = HonorKeepWithNextWhenBalancing,
        BalanceTableRowLines = BalanceTableRowLines,
        FinalColumnSpacingAfter = FinalColumnSpacingAfter,
        SeparatorColor = SeparatorColor,
        SeparatorWidth = SeparatorWidth
    };
}
