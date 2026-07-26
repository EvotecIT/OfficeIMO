namespace OfficeIMO.Html;

/// <summary>
/// Describes the deterministic device and user-preference values exposed to CSS media queries.
/// Geometry and screen/print media type remain owned by <see cref="HtmlRenderOptions"/>.
/// </summary>
public sealed class HtmlRenderMediaFeatures {
    /// <summary>Preferred light or dark color scheme. The default is light.</summary>
    public HtmlPreferredColorScheme PreferredColorScheme { get; set; } = HtmlPreferredColorScheme.Light;

    /// <summary>Preferred motion behavior. Static output defaults to no explicit reduction preference.</summary>
    public HtmlReducedMotionPreference ReducedMotion { get; set; } = HtmlReducedMotionPreference.NoPreference;

    /// <summary>Primary pointer capability. Static output defaults to no pointer.</summary>
    public HtmlPointerCapability Pointer { get; set; } = HtmlPointerCapability.None;

    /// <summary>Best available pointer capability. Static output defaults to no pointer.</summary>
    public HtmlPointerCapability AnyPointer { get; set; } = HtmlPointerCapability.None;

    /// <summary>Primary hover capability. Static output defaults to no hover state.</summary>
    public HtmlHoverCapability Hover { get; set; } = HtmlHoverCapability.None;

    /// <summary>Best available hover capability. Static output defaults to no hover state.</summary>
    public HtmlHoverCapability AnyHover { get; set; } = HtmlHoverCapability.None;

    /// <summary>Bits available per RGB color component. The default is 8.</summary>
    public int ColorBitsPerComponent { get; set; } = 8;

    /// <summary>Bits available per monochrome pixel. The default is zero for a color output.</summary>
    public int MonochromeBitsPerPixel { get; set; }

    /// <summary>CSS output resolution in dots per inch. The default is 96.</summary>
    public double ResolutionDpi { get; set; } = HtmlRenderOptions.CssPixelsPerInch;

    /// <summary>Creates an independent snapshot.</summary>
    public HtmlRenderMediaFeatures Clone() => new HtmlRenderMediaFeatures {
        PreferredColorScheme = PreferredColorScheme,
        ReducedMotion = ReducedMotion,
        Pointer = Pointer,
        AnyPointer = AnyPointer,
        Hover = Hover,
        AnyHover = AnyHover,
        ColorBitsPerComponent = ColorBitsPerComponent,
        MonochromeBitsPerPixel = MonochromeBitsPerPixel,
        ResolutionDpi = ResolutionDpi
    };

    internal void Validate() {
        ValidateEnum(PreferredColorScheme, nameof(PreferredColorScheme));
        ValidateEnum(ReducedMotion, nameof(ReducedMotion));
        ValidateEnum(Pointer, nameof(Pointer));
        ValidateEnum(AnyPointer, nameof(AnyPointer));
        ValidateEnum(Hover, nameof(Hover));
        ValidateEnum(AnyHover, nameof(AnyHover));
        if (ColorBitsPerComponent < 0) throw new ArgumentOutOfRangeException(nameof(ColorBitsPerComponent));
        if (MonochromeBitsPerPixel < 0) throw new ArgumentOutOfRangeException(nameof(MonochromeBitsPerPixel));
        if (ResolutionDpi <= 0D || double.IsNaN(ResolutionDpi) || double.IsInfinity(ResolutionDpi)) {
            throw new ArgumentOutOfRangeException(nameof(ResolutionDpi));
        }
    }

    private static void ValidateEnum<T>(T value, string parameterName) where T : struct {
        if (!Enum.IsDefined(typeof(T), value)) throw new ArgumentOutOfRangeException(parameterName);
    }
}
