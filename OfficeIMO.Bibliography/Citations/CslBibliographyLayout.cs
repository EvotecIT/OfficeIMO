using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

/// <summary>Alignment of subsequent lines relative to a bibliography entry's first field.</summary>
public enum CslSecondFieldAlignment {
    /// <summary>No automatic alignment after the first field.</summary>
    None,
    /// <summary>The first field starts at the margin; subsequent lines start at the second field.</summary>
    Flush,
    /// <summary>The first field sits outside the margin; subsequent lines start at the margin.</summary>
    Margin
}

/// <summary>Immutable CSL bibliography presentation settings for an output host.</summary>
/// <remarks>HTML entries contain CSL display classes. The host applies these spacing and indentation settings in its own layout system.</remarks>
public sealed class CslBibliographyLayout {
    internal CslBibliographyLayout(XElement bibliography, int maximumLeftMarginCharacters) {
        HangingIndent = (string?)bibliography.Attribute("hanging-indent") == "true";
        LineSpacing = ReadSpacing(bibliography, "line-spacing", 1);
        EntrySpacing = ReadSpacing(bibliography, "entry-spacing", 1);
        SecondFieldAlignment = (string?)bibliography.Attribute("second-field-align") switch {
            "flush" => CslSecondFieldAlignment.Flush,
            "margin" => CslSecondFieldAlignment.Margin,
            _ => CslSecondFieldAlignment.None
        };
        MaximumLeftMarginCharacters = maximumLeftMarginCharacters;
    }

    /// <summary>Whether the style requests a hanging indent.</summary>
    public bool HangingIndent { get; }
    /// <summary>Line-height multiplier. The CSL default is one.</summary>
    public int LineSpacing { get; }
    /// <summary>Additional line-height units between entries. The CSL default is one.</summary>
    public int EntrySpacing { get; }
    /// <summary>Alignment requested for lines after the first field.</summary>
    public CslSecondFieldAlignment SecondFieldAlignment { get; }
    /// <summary>Longest rendered left-margin field in UTF-16 characters across the returned entries.</summary>
    /// <remarks>This is a text-length hint, not a measured glyph width. A layout host may measure the actual fields.</remarks>
    public int MaximumLeftMarginCharacters { get; }

    private static int ReadSpacing(XElement bibliography, string name, int fallback) =>
        (string?)bibliography.Attribute(name) is string value ? int.Parse(value, CultureInfo.InvariantCulture) : fallback;
}
