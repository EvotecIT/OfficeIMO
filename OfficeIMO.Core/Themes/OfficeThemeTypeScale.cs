using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Document type sizes in physical units.</summary>
public sealed class OfficeThemeTypeScale {
    /// <summary>Gets the Title size.</summary>
    public OfficeLength Title { get; }
    /// <summary>Gets the Heading1 size.</summary>
    public OfficeLength Heading1 { get; }
    /// <summary>Gets the Heading2 size.</summary>
    public OfficeLength Heading2 { get; }
    /// <summary>Gets the Heading3 size.</summary>
    public OfficeLength Heading3 { get; }
    /// <summary>Gets the Heading4 size.</summary>
    public OfficeLength Heading4 { get; }
    /// <summary>Gets the Heading5 size.</summary>
    public OfficeLength Heading5 { get; }
    /// <summary>Gets the Heading6 size.</summary>
    public OfficeLength Heading6 { get; }
    /// <summary>Gets the Body size.</summary>
    public OfficeLength Body { get; }
    /// <summary>Gets the Caption size.</summary>
    public OfficeLength Caption { get; }

    /// <summary>Creates immutable document type sizes in physical units.</summary>
    public OfficeThemeTypeScale(OfficeLength title, OfficeLength heading1, OfficeLength heading2, OfficeLength heading3, OfficeLength heading4, OfficeLength heading5, OfficeLength heading6, OfficeLength body, OfficeLength caption) {
        Title = ThemeValues.NonNegative(title, nameof(title));
        Heading1 = ThemeValues.NonNegative(heading1, nameof(heading1));
        Heading2 = ThemeValues.NonNegative(heading2, nameof(heading2));
        Heading3 = ThemeValues.NonNegative(heading3, nameof(heading3));
        Heading4 = ThemeValues.NonNegative(heading4, nameof(heading4));
        Heading5 = ThemeValues.NonNegative(heading5, nameof(heading5));
        Heading6 = ThemeValues.NonNegative(heading6, nameof(heading6));
        Body = ThemeValues.NonNegative(body, nameof(body));
        Caption = ThemeValues.NonNegative(caption, nameof(caption));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeTypeScale With(OfficeLength? title = null, OfficeLength? heading1 = null, OfficeLength? heading2 = null, OfficeLength? heading3 = null, OfficeLength? heading4 = null, OfficeLength? heading5 = null, OfficeLength? heading6 = null, OfficeLength? body = null, OfficeLength? caption = null) =>
        new OfficeThemeTypeScale(title ?? Title, heading1 ?? Heading1, heading2 ?? Heading2, heading3 ?? Heading3, heading4 ?? Heading4, heading5 ?? Heading5, heading6 ?? Heading6, body ?? Body, caption ?? Caption);
}
