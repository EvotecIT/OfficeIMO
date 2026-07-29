using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one reusable, deterministic rendering configuration shared by OfficeIMO document adapters.
/// </summary>
/// <remarks>
/// A profile owns format-neutral resources and acceptance policy. Document-specific pagination and
/// layout settings remain on the corresponding Word, Excel, PowerPoint, PDF, Visio, or OneNote options.
/// </remarks>
public sealed class OfficeRenderingProfile {
    private readonly OfficeFontFaceCollection _fonts;
    private readonly OfficeImageExportPolicy _policy;

    /// <summary>Creates a rendering profile.</summary>
    /// <param name="name">Stable human-readable profile name used by diagnostics and host configuration.</param>
    /// <param name="fonts">Deterministic caller-supplied font faces and fallback order.</param>
    /// <param name="textShapingProvider">Optional complex-text shaping provider.</param>
    /// <param name="textShapingLanguage">Optional BCP 47 language hint.</param>
    /// <param name="imageCodec">Optional shared image codec.</param>
    /// <param name="policy">Diagnostic acceptance policy.</param>
    public OfficeRenderingProfile(
        string name,
        OfficeFontFaceCollection? fonts = null,
        IOfficeTextShapingProvider? textShapingProvider = null,
        string? textShapingLanguage = null,
        IOfficeRasterImageCodec? imageCodec = null,
        OfficeImageExportPolicy? policy = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("A rendering profile name is required.", nameof(name));
        }

        Name = name.Trim();
        _fonts = fonts?.Clone() ?? new OfficeFontFaceCollection();
        TextShapingProvider = textShapingProvider;
        TextShapingLanguage = string.IsNullOrWhiteSpace(textShapingLanguage)
            ? null
            : textShapingLanguage!.Trim();
        ImageCodec = imageCodec;
        _policy = policy?.Clone() ?? new OfficeImageExportPolicy();
    }

    /// <summary>
    /// Dependency-free profile using OfficeIMO's managed text shaper and default loss policy.
    /// </summary>
    public static OfficeRenderingProfile Managed { get; } = new OfficeRenderingProfile(
        "officeimo-managed",
        textShapingProvider: OfficeManagedTextShapingProvider.Instance);

    /// <summary>Stable human-readable profile name.</summary>
    public string Name { get; }

    /// <summary>Independent snapshot of deterministic font faces and fallback order.</summary>
    public OfficeFontFaceCollection Fonts => _fonts.Clone();

    /// <summary>Complex-text shaping provider used by all compatible renderers.</summary>
    public IOfficeTextShapingProvider? TextShapingProvider { get; }

    /// <summary>Optional BCP 47 language hint.</summary>
    public string? TextShapingLanguage { get; }

    /// <summary>Optional shared raster image codec.</summary>
    public IOfficeRasterImageCodec? ImageCodec { get; }

    /// <summary>Independent snapshot of the diagnostic acceptance policy.</summary>
    public OfficeImageExportPolicy Policy => _policy.Clone();

    internal OfficeFontFaceCollection FontsSnapshot => _fonts;

    internal OfficeImageExportPolicy PolicySnapshot => _policy;
}

/// <summary>Controls how a rendering profile is combined with existing export options.</summary>
public enum OfficeRenderingProfileApplyMode {
    /// <summary>Replace profile-owned fonts, shaping, codec, language, and policy settings.</summary>
    Replace,

    /// <summary>Merge fonts and apply only non-null optional profile services while replacing the policy.</summary>
    Overlay
}
