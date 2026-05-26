namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable heading styles for the built-in H1/H2/H3 levels.
/// </summary>
public sealed class PdfHeadingStyles {
    private PdfHeadingStyle? _level1;
    private PdfHeadingStyle? _level2;
    private PdfHeadingStyle? _level3;

    /// <summary>Default style for H1 headings.</summary>
    public PdfHeadingStyle? Level1 {
        get => _level1?.Clone();
        set => _level1 = value?.Clone();
    }

    /// <summary>Default style for H2 headings.</summary>
    public PdfHeadingStyle? Level2 {
        get => _level2?.Clone();
        set => _level2 = value?.Clone();
    }

    /// <summary>Default style for H3 headings.</summary>
    public PdfHeadingStyle? Level3 {
        get => _level3?.Clone();
        set => _level3 = value?.Clone();
    }

    /// <summary>Creates a deep copy of this heading style set.</summary>
    public PdfHeadingStyles Clone() {
        return new PdfHeadingStyles {
            Level1 = _level1?.Clone(),
            Level2 = _level2?.Clone(),
            Level3 = _level3?.Clone()
        };
    }

    internal PdfHeadingStyle? GetSnapshot(int level) {
        return level switch {
            1 => _level1,
            2 => _level2,
            3 => _level3,
            _ => null
        };
    }

    internal void Set(int level, PdfHeadingStyle style) {
        Guard.NotNull(style, nameof(style));
        switch (level) {
            case 1:
                _level1 = style.Clone();
                break;
            case 2:
                _level2 = style.Clone();
                break;
            case 3:
                _level3 = style.Clone();
                break;
            default:
                throw new System.ArgumentOutOfRangeException(nameof(level), "Heading level must be 1, 2, or 3.");
        }
    }
}
