namespace OfficeIMO.Pdf;

/// <summary>Optional accessibility attributes for a canvas structure container.</summary>
public sealed class PdfCanvasStructureOptions {
    private string? _alternativeText;
    private PdfCanvasTableHeaderScope? _headerScope;
    private int _columnSpan = 1;
    private int _rowSpan = 1;

    /// <summary>Alternative text associated with the structure container.</summary>
    public string? AlternativeText {
        get => _alternativeText;
        set {
            if (value != null) Guard.NotNullOrWhiteSpace(value, nameof(AlternativeText));
            _alternativeText = value?.Trim();
        }
    }

    /// <summary>Row, column, or combined scope for a table-header cell.</summary>
    public PdfCanvasTableHeaderScope? HeaderScope {
        get => _headerScope;
        set {
            if (value.HasValue && ((int)value.Value < (int)PdfCanvasTableHeaderScope.Row || (int)value.Value > (int)PdfCanvasTableHeaderScope.Both)) {
                throw new ArgumentOutOfRangeException(nameof(HeaderScope));
            }
            _headerScope = value;
        }
    }

    /// <summary>Number of table columns occupied by a tagged cell.</summary>
    public int ColumnSpan {
        get => _columnSpan;
        set => _columnSpan = ValidateSpan(value, nameof(ColumnSpan));
    }

    /// <summary>Number of table rows occupied by a tagged cell.</summary>
    public int RowSpan {
        get => _rowSpan;
        set => _rowSpan = ValidateSpan(value, nameof(RowSpan));
    }

    internal PdfCanvasStructureOptions Clone() => new PdfCanvasStructureOptions {
        AlternativeText = AlternativeText,
        HeaderScope = HeaderScope,
        ColumnSpan = ColumnSpan,
        RowSpan = RowSpan
    };

    private static int ValidateSpan(int value, string parameterName) {
        if (value < 1 || value > 1000) throw new ArgumentOutOfRangeException(parameterName, "Canvas table spans must be between 1 and 1000.");
        return value;
    }
}
