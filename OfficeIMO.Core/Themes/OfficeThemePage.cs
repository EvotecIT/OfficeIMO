using System;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Page margins for paged documents (Word and PDF), in physical units.</summary>
public sealed class OfficeThemePage {
    /// <summary>Gets the top margin.</summary>
    public OfficeLength MarginTop { get; }
    /// <summary>Gets the right margin.</summary>
    public OfficeLength MarginRight { get; }
    /// <summary>Gets the bottom margin.</summary>
    public OfficeLength MarginBottom { get; }
    /// <summary>Gets the left margin.</summary>
    public OfficeLength MarginLeft { get; }

    /// <summary>Creates immutable page margins.</summary>
    public OfficeThemePage(OfficeLength marginTop, OfficeLength marginRight, OfficeLength marginBottom, OfficeLength marginLeft) {
        MarginTop = ThemeValues.NonNegative(marginTop, nameof(marginTop));
        MarginRight = ThemeValues.NonNegative(marginRight, nameof(marginRight));
        MarginBottom = ThemeValues.NonNegative(marginBottom, nameof(marginBottom));
        MarginLeft = ThemeValues.NonNegative(marginLeft, nameof(marginLeft));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemePage With(OfficeLength? marginTop = null, OfficeLength? marginRight = null, OfficeLength? marginBottom = null, OfficeLength? marginLeft = null) =>
        new OfficeThemePage(marginTop ?? MarginTop, marginRight ?? MarginRight, marginBottom ?? MarginBottom, marginLeft ?? MarginLeft);
}
