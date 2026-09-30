namespace OfficeIMO.Rtf;

/// <summary>
/// Border formatting for one side of an RTF paragraph.
/// </summary>
public sealed partial class RtfParagraphBorder {
    /// <summary>Border line style.</summary>
    public RtfParagraphBorderStyle Style { get => DirectStyle ?? RtfParagraphBorderStyle.None; set => DirectStyle = value; }

    /// <summary>Authored border style. An explicit None clears an inherited border; null inherits.</summary>
    public RtfParagraphBorderStyle? DirectStyle { get; set; }

    /// <summary>Border width value carried by the RTF <c>\brdrw</c> control.</summary>
    public int? Width { get; set; }

    /// <summary>One-based color table index.</summary>
    public int? ColorIndex { get; set; }

    /// <summary>Whether any border formatting is present.</summary>
    public bool HasAnyValue =>
        DirectStyle.HasValue ||
        Width.HasValue ||
        ColorIndex.HasValue;
}
