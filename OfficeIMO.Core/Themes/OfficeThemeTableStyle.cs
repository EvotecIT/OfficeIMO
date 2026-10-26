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

    /// <summary>Creates immutable table colors and geometry shared by authoring formats.</summary>
    public OfficeThemeTableStyle(OfficeColor headerFill, OfficeColor headerText, bool banding, OfficeColor bandFill, OfficeColor borderColor, OfficeLength borderWidth, OfficeLength cellPadding) {
        HeaderFill = headerFill;
        HeaderText = headerText;
        Banding = banding;
        BandFill = bandFill;
        BorderColor = borderColor;
        BorderWidth = ThemeValues.NonNegative(borderWidth, nameof(borderWidth));
        CellPadding = ThemeValues.NonNegative(cellPadding, nameof(cellPadding));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeTableStyle With(OfficeColor? headerFill = null, OfficeColor? headerText = null, bool? banding = null, OfficeColor? bandFill = null, OfficeColor? borderColor = null, OfficeLength? borderWidth = null, OfficeLength? cellPadding = null) =>
        new OfficeThemeTableStyle(headerFill ?? HeaderFill, headerText ?? HeaderText, banding ?? Banding, bandFill ?? BandFill, borderColor ?? BorderColor, borderWidth ?? BorderWidth, cellPadding ?? CellPadding);
}
