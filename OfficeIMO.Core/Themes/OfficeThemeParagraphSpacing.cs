using System;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Space before and after one paragraph level, in physical units.</summary>
public sealed class OfficeThemeParagraphSpacing {
    /// <summary>Gets the space before the paragraph.</summary>
    public OfficeLength Before { get; }
    /// <summary>Gets the space after the paragraph.</summary>
    public OfficeLength After { get; }

    /// <summary>Creates immutable paragraph spacing.</summary>
    public OfficeThemeParagraphSpacing(OfficeLength before, OfficeLength after) {
        Before = ThemeValues.NonNegative(before, nameof(before));
        After = ThemeValues.NonNegative(after, nameof(after));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeParagraphSpacing With(OfficeLength? before = null, OfficeLength? after = null) =>
        new OfficeThemeParagraphSpacing(before ?? Before, after ?? After);
}
