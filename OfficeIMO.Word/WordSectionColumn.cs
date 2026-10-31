namespace OfficeIMO.Word;

/// <summary>Describes an explicitly sized section column using physical lengths.</summary>
public sealed class WordSectionColumn {
    /// <summary>Creates an immutable column width and optional following gap.</summary>
    /// <param name="width">Column width that rounds to a positive 32-bit twip value.</param>
    /// <param name="spaceAfter">Optional non-negative gap that rounds to a 32-bit twip value. Omission means zero for unequal columns.</param>
    /// <exception cref="ArgumentOutOfRangeException">The width rounds to zero or is negative, or the gap is negative.</exception>
    /// <exception cref="OverflowException">A rounded dimension exceeds the 32-bit twip range.</exception>
    public WordSectionColumn(OfficeLength width, OfficeLength? spaceAfter = null) {
        long widthTwips = width.ToTwips();
        if (widthTwips <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        _ = checked((int)widthTwips);
        if (spaceAfter.HasValue) {
            if (spaceAfter.Value.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(spaceAfter));
            _ = checked((int)spaceAfter.Value.ToTwips());
        }
        Width = width;
        SpaceAfter = spaceAfter;
    }

    /// <summary>Gets the explicitly authored column width. Storage rounds to whole twips.</summary>
    public OfficeLength Width { get; }
    /// <summary>Gets the optional gap following this column. Null preserves omission in DOCX; native DOC writes its effective zero.</summary>
    public OfficeLength? SpaceAfter { get; }
}
