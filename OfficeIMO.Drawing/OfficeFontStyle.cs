using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes text style flags without depending on a rendering or font library.
/// </summary>
[Flags]
public enum OfficeFontStyle {
    /// <summary>Regular text.</summary>
    Regular = 0,

    /// <summary>Bold text.</summary>
    Bold = 1,

    /// <summary>Italic text.</summary>
    Italic = 2,

    /// <summary>Underlined text.</summary>
    Underline = 4
}
