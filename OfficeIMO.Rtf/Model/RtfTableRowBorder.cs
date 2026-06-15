namespace OfficeIMO.Rtf;

/// <summary>
/// Border formatting for one side of an RTF table row.
/// </summary>
public sealed class RtfTableRowBorder {
    /// <summary>Border line style.</summary>
    public RtfTableCellBorderStyle Style { get; set; } = RtfTableCellBorderStyle.None;

    /// <summary>Border width value carried by the RTF <c>\brdrw</c> control.</summary>
    public int? Width { get; set; }

    /// <summary>One-based color table index.</summary>
    public int? ColorIndex { get; set; }

    /// <summary>Whether any border formatting is present.</summary>
    public bool HasAnyValue =>
        Style != RtfTableCellBorderStyle.None ||
        Width.HasValue ||
        ColorIndex.HasValue;
}
