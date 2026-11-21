namespace OfficeIMO.Word;

/// <summary>Options controlling Word accessibility inspection.</summary>
public sealed class WordAccessibilityOptions {
    private double _minimumTextContrastRatio = 4.5D;
    private double _minimumLargeTextContrastRatio = 3D;

    /// <summary>Selected policy profile.</summary>
    public OfficeAccessibilityPolicyProfile Profile { get; set; } = OfficeAccessibilityPolicyProfile.Default;
    /// <summary>Whether the document must define a title. Defaults to true.</summary>
    public bool RequireDocumentTitle { get; set; } = true;
    /// <summary>Whether the document must define a default text language. Defaults to true.</summary>
    public bool RequireLanguage { get; set; } = true;
    /// <summary>Whether images, shapes, charts and SmartArt need alternative text unless marked decorative. Defaults to true.</summary>
    public bool RequireAlternativeText { get; set; } = true;
    /// <summary>Whether data tables need a header row. Defaults to true.</summary>
    public bool RequireTableHeaders { get; set; } = true;
    /// <summary>Whether heading levels are checked for skips, and long documents for missing headings. Defaults to true.</summary>
    public bool CheckHeadings { get; set; } = true;
    /// <summary>Whether hyperlinks need meaningful visible text. Defaults to true.</summary>
    public bool CheckMeaningfulLinks { get; set; } = true;
    /// <summary>Whether text with an explicit colour is checked against its shading or page background. Defaults to true.</summary>
    public bool CheckContrast { get; set; } = true;
    /// <summary>Whether objects that are not inline with text are reported. Defaults to true.</summary>
    public bool CheckFloatingObjects { get; set; } = true;
    /// <summary>Whether consecutive blank paragraphs used for spacing are reported. Defaults to true.</summary>
    public bool CheckBlankParagraphs { get; set; } = true;

    /// <summary>Minimum contrast ratio for normal text. Defaults to the WCAG AA value 4.5.</summary>
    public double MinimumTextContrastRatio {
        get => _minimumTextContrastRatio;
        set => _minimumTextContrastRatio = ValidateContrast(value, nameof(MinimumTextContrastRatio));
    }

    /// <summary>Minimum contrast ratio for large text (18 pt, or 14 pt bold). Defaults to the WCAG AA value 3.</summary>
    public double MinimumLargeTextContrastRatio {
        get => _minimumLargeTextContrastRatio;
        set => _minimumLargeTextContrastRatio = ValidateContrast(value, nameof(MinimumLargeTextContrastRatio));
    }

    private static double ValidateContrast(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 1D || value > 21D) {
            throw new ArgumentOutOfRangeException(name, "Contrast ratio must be between 1 and 21.");
        }
        return value;
    }
}

/// <summary>Safe automatic fixes that <see cref="WordDocument.FixAccessibility"/> applies.</summary>
public sealed class WordAccessibilityFixOptions {
    /// <summary>Title to set when the document has none.</summary>
    public string? Title { get; set; }
    /// <summary>BCP 47 language, such as <c>en-US</c>, to set as the document default when none is defined.</summary>
    public string? Language { get; set; }
    /// <summary>Whether to mark the first row of data tables without a header row as the header row. Tables that look like layout tables are left alone.</summary>
    public bool MarkTableHeaderRows { get; set; }
}
