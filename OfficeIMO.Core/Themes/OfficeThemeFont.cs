using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>A preferred font family with an ordered, immutable fallback list.</summary>
public sealed class OfficeThemeFont {
    /// <summary>Gets the preferred font family.</summary>
    public string Family { get; }
    /// <summary>Gets fallback families in preference order.</summary>
    public IReadOnlyList<string> Fallbacks { get; }

    /// <summary>Creates immutable a preferred font family with an ordered, immutable fallback list.</summary>
    public OfficeThemeFont(string family, IReadOnlyList<string> fallbacks) {
        Family = ThemeValues.FontName(family, nameof(family));
        Fallbacks = ThemeValues.Names(fallbacks, nameof(fallbacks));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeFont With(string? family = null, IReadOnlyList<string>? fallbacks = null) =>
        new OfficeThemeFont(family ?? Family, fallbacks ?? Fallbacks);
}
