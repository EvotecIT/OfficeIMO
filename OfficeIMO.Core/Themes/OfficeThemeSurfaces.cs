using System;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Background tints and structural colours: panels, hairlines, the accent tint and status row fills.</summary>
public sealed class OfficeThemeSurfaces {
    /// <summary>Gets the fill for panels, callouts and code blocks.</summary>
    public OfficeColor Surface { get; }
    /// <summary>Gets the colour of hairlines and separators.</summary>
    public OfficeColor Rule { get; }
    /// <summary>Gets <see cref="OfficeColor.Accent1"/> at 10 % over the background.</summary>
    public OfficeColor Accent1Tint { get; }
    /// <summary>Gets the row fill behind a success status.</summary>
    public OfficeColor SuccessFill { get; }
    /// <summary>Gets the row fill behind a warning status.</summary>
    public OfficeColor WarningFill { get; }
    /// <summary>Gets the row fill behind a danger status.</summary>
    public OfficeColor DangerFill { get; }

    /// <summary>Creates immutable surface colours.</summary>
    public OfficeThemeSurfaces(OfficeColor surface, OfficeColor rule, OfficeColor accent1Tint, OfficeColor successFill, OfficeColor warningFill, OfficeColor dangerFill) {
        Surface = surface;
        Rule = rule;
        Accent1Tint = accent1Tint;
        SuccessFill = successFill;
        WarningFill = warningFill;
        DangerFill = dangerFill;
    }

    /// <summary>
    /// Derives the accent and status tints from a palette: Accent1 at 10 % and each status colour at 12 % (18 % on a dark
    /// background) composited over the background.
    /// </summary>
    public static OfficeThemeSurfaces FromPalette(OfficeThemePalette palette, OfficeColor surface, OfficeColor rule) {
        if (palette == null) throw new ArgumentNullException(nameof(palette));
        OfficeColor background = palette.Background;
        double status = IsDark(background) ? 0.18 : 0.12;
        return new OfficeThemeSurfaces(surface, rule, Tint(palette.Accent1, background, 0.10),
            Tint(palette.Success, background, status), Tint(palette.Warning, background, status), Tint(palette.Danger, background, status));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeSurfaces With(OfficeColor? surface = null, OfficeColor? rule = null, OfficeColor? accent1Tint = null,
        OfficeColor? successFill = null, OfficeColor? warningFill = null, OfficeColor? dangerFill = null) =>
        new OfficeThemeSurfaces(surface ?? Surface, rule ?? Rule, accent1Tint ?? Accent1Tint, successFill ?? SuccessFill,
            warningFill ?? WarningFill, dangerFill ?? DangerFill);

    internal static bool IsDark(OfficeColor background) =>
        0.2126 * background.R + 0.7152 * background.G + 0.0722 * background.B < 128;

    private static OfficeColor Tint(OfficeColor color, OfficeColor background, double amount) =>
        OfficeColor.FromRgb(Mix(color.R, background.R, amount), Mix(color.G, background.G, amount), Mix(color.B, background.B, amount));

    private static byte Mix(byte color, byte background, double amount) =>
        (byte)Math.Round(background + (color - background) * amount, MidpointRounding.AwayFromZero);
}
