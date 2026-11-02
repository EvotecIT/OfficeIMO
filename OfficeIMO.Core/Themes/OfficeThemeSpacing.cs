using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Paragraph, list and slide spacing in physical units.</summary>
public sealed class OfficeThemeSpacing {
    /// <summary>Gets space before paragraphs.</summary>
    public OfficeLength ParagraphBefore { get; }
    /// <summary>Gets space after paragraphs.</summary>
    public OfficeLength ParagraphAfter { get; }
    /// <summary>Gets list indentation.</summary>
    public OfficeLength ListIndent { get; }
    /// <summary>Gets the top slide margin.</summary>
    public OfficeLength SlideMarginTop { get; }
    /// <summary>Gets the right slide margin.</summary>
    public OfficeLength SlideMarginRight { get; }
    /// <summary>Gets the bottom slide margin.</summary>
    public OfficeLength SlideMarginBottom { get; }
    /// <summary>Gets the left slide margin.</summary>
    public OfficeLength SlideMarginLeft { get; }

    /// <summary>Creates immutable paragraph, list and slide spacing in physical units.</summary>
    public OfficeThemeSpacing(OfficeLength paragraphBefore, OfficeLength paragraphAfter, OfficeLength listIndent, OfficeLength slideMarginTop, OfficeLength slideMarginRight, OfficeLength slideMarginBottom, OfficeLength slideMarginLeft) {
        ParagraphBefore = ThemeValues.NonNegative(paragraphBefore, nameof(paragraphBefore));
        ParagraphAfter = ThemeValues.NonNegative(paragraphAfter, nameof(paragraphAfter));
        ListIndent = ThemeValues.NonNegative(listIndent, nameof(listIndent));
        SlideMarginTop = ThemeValues.NonNegative(slideMarginTop, nameof(slideMarginTop));
        SlideMarginRight = ThemeValues.NonNegative(slideMarginRight, nameof(slideMarginRight));
        SlideMarginBottom = ThemeValues.NonNegative(slideMarginBottom, nameof(slideMarginBottom));
        SlideMarginLeft = ThemeValues.NonNegative(slideMarginLeft, nameof(slideMarginLeft));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeSpacing With(OfficeLength? paragraphBefore = null, OfficeLength? paragraphAfter = null, OfficeLength? listIndent = null, OfficeLength? slideMarginTop = null, OfficeLength? slideMarginRight = null, OfficeLength? slideMarginBottom = null, OfficeLength? slideMarginLeft = null) =>
        new OfficeThemeSpacing(paragraphBefore ?? ParagraphBefore, paragraphAfter ?? ParagraphAfter, listIndent ?? ListIndent, slideMarginTop ?? SlideMarginTop, slideMarginRight ?? SlideMarginRight, slideMarginBottom ?? SlideMarginBottom, slideMarginLeft ?? SlideMarginLeft);
}
