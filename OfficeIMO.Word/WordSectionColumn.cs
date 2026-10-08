namespace OfficeIMO.Word;

/// <summary>Describes an explicitly sized section column in twips (one twentieth of a point).</summary>
public sealed class WordSectionColumn {
    /// <summary>Creates an immutable column width and optional following gap.</summary>
    /// <param name="widthTwips">Positive column width in twips.</param>
    /// <param name="spaceAfterTwips">Optional non-negative gap after this column in twips.</param>
    public WordSectionColumn(int widthTwips, int? spaceAfterTwips = null) {
        if (widthTwips <= 0) throw new ArgumentOutOfRangeException(nameof(widthTwips));
        if (spaceAfterTwips < 0) throw new ArgumentOutOfRangeException(nameof(spaceAfterTwips));
        WidthTwips = widthTwips;
        SpaceAfterTwips = spaceAfterTwips;
    }

    /// <summary>Gets the explicitly authored column width in twips.</summary>
    public int WidthTwips { get; }
    /// <summary>Gets the optional gap following this column in twips.</summary>
    public int? SpaceAfterTwips { get; }
}
