using System;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Text colours, line spacing and heading treatment shared by document formats.</summary>
public sealed class OfficeThemeTextStyle {
    /// <summary>Gets the line spacing multiple for body text (1.0 is single spacing).</summary>
    public double LineSpacing { get; }
    /// <summary>Gets the title colour.</summary>
    public OfficeColor TitleColor { get; }
    /// <summary>Gets the subtitle colour.</summary>
    public OfficeColor SubtitleColor { get; }
    /// <summary>Gets the colour of heading levels 1 to 6.</summary>
    public OfficeColor HeadingColor { get; }
    /// <summary>Gets the hyperlink colour.</summary>
    public OfficeColor LinkColor { get; }
    /// <summary>Gets whether headings are bold.</summary>
    public bool HeadingBold { get; }
    /// <summary>Gets the rule drawn under the title, or <see cref="OfficeThemeRule.None"/>.</summary>
    public OfficeThemeRule TitleRule { get; }
    /// <summary>Gets the space before and after the title.</summary>
    public OfficeThemeParagraphSpacing TitleSpacing { get; }
    /// <summary>Gets the space before and after level 1 headings.</summary>
    public OfficeThemeParagraphSpacing Heading1Spacing { get; }
    /// <summary>Gets the space before and after level 2 headings.</summary>
    public OfficeThemeParagraphSpacing Heading2Spacing { get; }
    /// <summary>Gets the space before and after level 3 to 6 headings.</summary>
    public OfficeThemeParagraphSpacing Heading3Spacing { get; }

    /// <summary>Creates immutable text colours, line spacing and heading treatment.</summary>
    public OfficeThemeTextStyle(double lineSpacing, OfficeColor titleColor, OfficeColor subtitleColor, OfficeColor headingColor,
        OfficeColor linkColor, bool headingBold, OfficeThemeRule titleRule, OfficeThemeParagraphSpacing titleSpacing,
        OfficeThemeParagraphSpacing heading1Spacing, OfficeThemeParagraphSpacing heading2Spacing, OfficeThemeParagraphSpacing heading3Spacing) {
        if (double.IsNaN(lineSpacing) || double.IsInfinity(lineSpacing) || lineSpacing <= 0)
            throw new ArgumentOutOfRangeException(nameof(lineSpacing), "Line spacing must be a positive multiple.");
        LineSpacing = lineSpacing;
        TitleColor = titleColor;
        SubtitleColor = subtitleColor;
        HeadingColor = headingColor;
        LinkColor = linkColor;
        HeadingBold = headingBold;
        TitleRule = titleRule ?? throw new ArgumentNullException(nameof(titleRule));
        TitleSpacing = titleSpacing ?? throw new ArgumentNullException(nameof(titleSpacing));
        Heading1Spacing = heading1Spacing ?? throw new ArgumentNullException(nameof(heading1Spacing));
        Heading2Spacing = heading2Spacing ?? throw new ArgumentNullException(nameof(heading2Spacing));
        Heading3Spacing = heading3Spacing ?? throw new ArgumentNullException(nameof(heading3Spacing));
    }

    /// <summary>Gets the spacing for a heading level: 0 is the title, 1 to 6 are heading levels; levels 4 to 6 share level 3 spacing.</summary>
    public OfficeThemeParagraphSpacing GetHeadingSpacing(int level) => level switch {
        0 => TitleSpacing,
        1 => Heading1Spacing,
        2 => Heading2Spacing,
        >= 3 and <= 6 => Heading3Spacing,
        _ => throw new ArgumentOutOfRangeException(nameof(level), "Heading levels are 0 (title) to 6.")
    };

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeTextStyle With(double? lineSpacing = null, OfficeColor? titleColor = null, OfficeColor? subtitleColor = null,
        OfficeColor? headingColor = null, OfficeColor? linkColor = null, bool? headingBold = null, OfficeThemeRule? titleRule = null,
        OfficeThemeParagraphSpacing? titleSpacing = null, OfficeThemeParagraphSpacing? heading1Spacing = null,
        OfficeThemeParagraphSpacing? heading2Spacing = null, OfficeThemeParagraphSpacing? heading3Spacing = null) =>
        new OfficeThemeTextStyle(lineSpacing ?? LineSpacing, titleColor ?? TitleColor, subtitleColor ?? SubtitleColor,
            headingColor ?? HeadingColor, linkColor ?? LinkColor, headingBold ?? HeadingBold, titleRule ?? TitleRule,
            titleSpacing ?? TitleSpacing, heading1Spacing ?? Heading1Spacing, heading2Spacing ?? Heading2Spacing, heading3Spacing ?? Heading3Spacing);
}
