using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Reusable font fallback-pack activation for HTML renderers.</summary>
public static class HtmlRenderFontFallbackPackExtensions {
    /// <summary>
    /// Applies an immutable fallback pack and its default family list to HTML image or PDF options.
    /// </summary>
    public static TOptions UseFontFallbackPack<TOptions>(
        this TOptions options,
        OfficeFontFallbackPack pack,
        OfficeRenderingProfileApplyMode mode = OfficeRenderingProfileApplyMode.Replace)
        where TOptions : HtmlRenderOptions {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(pack);
        options.UseRenderingProfile(pack.CreateRenderingProfile(), mode);
        options.DefaultFontFamily = pack.DefaultFamilyNames;
        return options;
    }
}
