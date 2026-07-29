using OfficeIMO.Drawing;

namespace OfficeIMO.Drawing.HarfBuzz;

/// <summary>Creates reusable rendering profiles backed by the first-party HarfBuzz adapter.</summary>
public static class OfficeHarfBuzzRenderingProfile {
    /// <summary>Default HarfBuzz profile without caller-supplied fonts or language policy.</summary>
    public static OfficeRenderingProfile Default { get; } = Create();

    /// <summary>
    /// Creates a deterministic profile that routes complex-script shaping through HarfBuzz.
    /// </summary>
    /// <param name="fonts">Optional embedded font collection and fallback order.</param>
    /// <param name="language">Optional BCP 47 shaping language.</param>
    /// <param name="imageCodec">Optional shared image codec.</param>
    /// <param name="policy">Optional diagnostic acceptance policy.</param>
    public static OfficeRenderingProfile Create(
        OfficeFontFaceCollection? fonts = null,
        string? language = null,
        IOfficeRasterImageCodec? imageCodec = null,
        OfficeImageExportPolicy? policy = null) =>
        new OfficeRenderingProfile(
            "officeimo-harfbuzz",
            fonts,
            OfficeHarfBuzzTextShapingProvider.Instance,
            language,
            imageCodec,
            policy);
}
