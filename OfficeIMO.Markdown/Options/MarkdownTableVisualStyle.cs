namespace OfficeIMO.Markdown;

/// <summary>
/// Renderer-neutral table styling hints for Markdown exporters.
/// </summary>
/// <remarks>
/// Defaults retain their point magnitudes in Word and PDF. HTML converts physical lengths
/// to CSS pixels at 96 DPI; use <see cref="OfficeLength.Pixels"/> for an exact CSS pixel size.
/// </remarks>
public sealed class MarkdownTableVisualStyle {
    private OfficeLength _borderWidth = OfficeLength.Points(0.6);
    private OfficeLength _cellPaddingX = OfficeLength.Points(6);
    private OfficeLength _cellPaddingY = OfficeLength.Points(5);

    /// <summary>When true, exporters should use alternating row fills where supported.</summary>
    public bool UseRowStripes { get; set; } = true;

    /// <summary>When true, exporters should emphasize the first row as a header.</summary>
    public bool EmphasizeHeader { get; set; } = true;

    /// <summary>MarkdownTable border width in physical lengths.</summary>
    public OfficeLength BorderWidth {
        get => _borderWidth;
        set => _borderWidth = ValidateNonNegative(value, nameof(BorderWidth));
    }

    /// <summary>Horizontal cell padding in physical lengths.</summary>
    public OfficeLength CellPaddingX {
        get => _cellPaddingX;
        set => _cellPaddingX = ValidateNonNegative(value, nameof(CellPaddingX));
    }

    /// <summary>Vertical cell padding in physical lengths.</summary>
    public OfficeLength CellPaddingY {
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

    private static OfficeLength ValidateNonNegative(OfficeLength value, string name) {
        if (value < OfficeLength.Zero) {
            throw new ArgumentOutOfRangeException(name, "The value must be a finite non-negative number.");
        }

        return value;
    }
}
