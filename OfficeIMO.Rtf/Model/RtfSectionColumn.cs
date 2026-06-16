namespace OfficeIMO.Rtf;

/// <summary>
/// Section column definition used when an RTF section has unequal column widths or gaps.
/// </summary>
public sealed class RtfSectionColumn {
    /// <summary>Creates an empty section column definition.</summary>
    public RtfSectionColumn() {
    }

    /// <summary>Creates a section column definition with optional width and spacing.</summary>
    public RtfSectionColumn(int? widthTwips = null, int? spaceAfterTwips = null) {
        WidthTwips = widthTwips;
        SpaceAfterTwips = spaceAfterTwips;
    }

    /// <summary>Column width in twips, represented by <c>\colw</c>.</summary>
    public int? WidthTwips { get; set; }

    /// <summary>Space after the column in twips, represented by <c>\colsr</c>.</summary>
    public int? SpaceAfterTwips { get; set; }

    internal bool HasAnyValue => WidthTwips.HasValue || SpaceAfterTwips.HasValue;
}
