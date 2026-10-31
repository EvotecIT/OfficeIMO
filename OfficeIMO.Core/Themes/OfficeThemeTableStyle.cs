using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Table colors and geometry shared by authoring formats.</summary>
public sealed class OfficeThemeTableStyle {
    /// <summary>Gets the header background.</summary>
    public OfficeColor HeaderFill { get; }
    /// <summary>Gets the header text color.</summary>
    public OfficeColor HeaderText { get; }
    /// <summary>Gets whether alternating row fills are enabled.</summary>
    public bool Banding { get; }
    /// <summary>Gets the alternating row fill.</summary>
    public OfficeColor BandFill { get; }
    /// <summary>Gets table border color.</summary>
    public OfficeColor BorderColor { get; }
    /// <summary>Gets table border width.</summary>
    public OfficeLength BorderWidth { get; }
    /// <summary>Gets padding on each cell edge.</summary>
    public OfficeLength CellPadding { get; }
    /// <summary>Gets which borders are drawn.</summary>
    public OfficeThemeTableBorderMode BorderMode { get; }
    /// <summary>Gets the rule under the header row, or <see cref="OfficeThemeRule.None"/>.</summary>
    public OfficeThemeRule HeaderRule { get; }
    /// <summary>Gets whether header text is bold.</summary>
    public bool HeaderBold { get; }
    /// <summary>Gets the alignment of numeric cells.</summary>
    public OfficeAlignment NumberAlignment { get; }

    /// <summary>Creates immutable table colors and geometry with horizontal borders, a bold header, no header rule and right-aligned numbers.</summary>
    public OfficeThemeTableStyle(OfficeColor headerFill, OfficeColor headerText, bool banding, OfficeColor bandFill, OfficeColor borderColor, OfficeLength borderWidth, OfficeLength cellPadding)
        : this(headerFill, headerText, banding, bandFill, borderColor, borderWidth, cellPadding,
            OfficeThemeTableBorderMode.Horizontal, OfficeThemeRule.None, true, OfficeAlignment.Right) {
    }

    /// <summary>Creates immutable table colors, geometry, border mode, header treatment and numeric alignment.</summary>
    public OfficeThemeTableStyle(OfficeColor headerFill, OfficeColor headerText, bool banding, OfficeColor bandFill, OfficeColor borderColor, OfficeLength borderWidth, OfficeLength cellPadding,
        OfficeThemeTableBorderMode borderMode, OfficeThemeRule headerRule, bool headerBold, OfficeAlignment numberAlignment) {
        HeaderFill = headerFill;
        HeaderText = headerText;
        Banding = banding;
        BandFill = bandFill;
        BorderColor = borderColor;
        BorderWidth = ThemeValues.NonNegative(borderWidth, nameof(borderWidth));
        CellPadding = ThemeValues.NonNegative(cellPadding, nameof(cellPadding));
        BorderMode = borderMode;
        HeaderRule = headerRule ?? throw new ArgumentNullException(nameof(headerRule));
        HeaderBold = headerBold;
        NumberAlignment = numberAlignment;
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeTableStyle With(OfficeColor? headerFill = null, OfficeColor? headerText = null, bool? banding = null, OfficeColor? bandFill = null, OfficeColor? borderColor = null, OfficeLength? borderWidth = null, OfficeLength? cellPadding = null,
        OfficeThemeTableBorderMode? borderMode = null, OfficeThemeRule? headerRule = null, bool? headerBold = null, OfficeAlignment? numberAlignment = null) =>
        new OfficeThemeTableStyle(headerFill ?? HeaderFill, headerText ?? HeaderText, banding ?? Banding, bandFill ?? BandFill, borderColor ?? BorderColor, borderWidth ?? BorderWidth, cellPadding ?? CellPadding,
            borderMode ?? BorderMode, headerRule ?? HeaderRule, headerBold ?? HeaderBold, numberAlignment ?? NumberAlignment);
}
