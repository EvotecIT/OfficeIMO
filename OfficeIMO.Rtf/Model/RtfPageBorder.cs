namespace OfficeIMO.Rtf;

/// <summary>
/// Formatting for one side of an RTF page border.
/// </summary>
public sealed class RtfPageBorder {
    /// <summary>Border line style.</summary>
    public RtfPageBorderStyle Style { get; set; } = RtfPageBorderStyle.None;

    /// <summary>Border width value carried by the RTF <c>\brdrw</c> control.</summary>
    public int? Width { get; set; }

    /// <summary>Space between the page border and the measured edge, represented by <c>\brsp</c>.</summary>
    public int? Space { get; set; }

    /// <summary>One-based color table index.</summary>
    public int? ColorIndex { get; set; }

    /// <summary>Whether the border has a shadow effect.</summary>
    public bool Shadow { get; set; }

    /// <summary>Whether the border is displayed as a frame border.</summary>
    public bool Frame { get; set; }

    /// <summary>Sets the page border line properties.</summary>
    public RtfPageBorder Set(RtfPageBorderStyle style, int? width = null, int? space = null, int? colorIndex = null) {
        ValidateNonNegative(width, nameof(width));
        ValidateNonNegative(space, nameof(space));
        ValidateNonNegative(colorIndex, nameof(colorIndex));
        Style = style;
        Width = width;
        Space = space;
        ColorIndex = colorIndex;
        return this;
    }

    /// <summary>Whether any border formatting is present.</summary>
    public bool HasAnyValue =>
        Style != RtfPageBorderStyle.None ||
        Width.HasValue ||
        Space.HasValue ||
        ColorIndex.HasValue ||
        Shadow ||
        Frame;

    internal void Clear() {
        Style = RtfPageBorderStyle.None;
        Width = null;
        Space = null;
        ColorIndex = null;
        Shadow = false;
        Frame = false;
    }

    private static void ValidateNonNegative(int? value, string parameterName) {
        if (value.HasValue && value.Value < 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Page border value cannot be negative.");
        }
    }
}
