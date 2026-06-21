namespace OfficeIMO.Markdown;

/// <summary>
/// Renderer-neutral table styling hints for Markdown exporters.
/// </summary>
public sealed class MarkdownTableVisualStyle {
    private double _borderWidth = 0.6;
    private double _cellPaddingX = 6;
    private double _cellPaddingY = 5;

    /// <summary>When true, exporters should use alternating row fills where supported.</summary>
    public bool UseRowStripes { get; set; } = true;

    /// <summary>When true, exporters should emphasize the first row as a header.</summary>
    public bool EmphasizeHeader { get; set; } = true;

    /// <summary>Table border width in renderer points/pixels where applicable.</summary>
    public double BorderWidth {
        get => _borderWidth;
        set => _borderWidth = ValidateNonNegative(value, nameof(BorderWidth));
    }

    /// <summary>Horizontal cell padding in renderer points/pixels where applicable.</summary>
    public double CellPaddingX {
        get => _cellPaddingX;
        set => _cellPaddingX = ValidateNonNegative(value, nameof(CellPaddingX));
    }

    /// <summary>Vertical cell padding in renderer points/pixels where applicable.</summary>
    public double CellPaddingY {
        get => _cellPaddingY;
        set => _cellPaddingY = ValidateNonNegative(value, nameof(CellPaddingY));
    }

    /// <summary>Creates a copy of this table style.</summary>
    public MarkdownTableVisualStyle Clone() => new MarkdownTableVisualStyle {
        UseRowStripes = UseRowStripes,
        EmphasizeHeader = EmphasizeHeader,
        BorderWidth = BorderWidth,
        CellPaddingX = CellPaddingX,
        CellPaddingY = CellPaddingY
    };

    private static double ValidateNonNegative(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) {
            throw new ArgumentOutOfRangeException(name, "The value must be a finite non-negative number.");
        }

        return value;
    }
}
