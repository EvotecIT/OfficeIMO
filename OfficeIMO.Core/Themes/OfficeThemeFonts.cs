using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Heading and body font choices.</summary>
public sealed class OfficeThemeFonts {
    /// <summary>Gets heading fonts.</summary>
    public OfficeThemeFont Heading { get; }
    /// <summary>Gets body fonts.</summary>
    public OfficeThemeFont Body { get; }

    /// <summary>Creates immutable heading and body font choices.</summary>
    public OfficeThemeFonts(OfficeThemeFont heading, OfficeThemeFont body) {
        Heading = heading ?? throw new ArgumentNullException(nameof(heading));
        Body = body ?? throw new ArgumentNullException(nameof(body));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeFonts With(OfficeThemeFont? heading = null, OfficeThemeFont? body = null) =>
        new OfficeThemeFonts(heading ?? Heading, body ?? Body);
}
