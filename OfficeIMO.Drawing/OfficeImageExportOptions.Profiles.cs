using System;

namespace OfficeIMO.Drawing;

/// <summary>Shared rendering-profile configuration for all image export option types.</summary>
public static class OfficeImageExportOptionsExtensions {
    /// <summary>
    /// Applies one shared rendering profile to these options.
    /// </summary>
    /// <param name="options">Image-export options to configure.</param>
    /// <param name="profile">Reusable format-neutral rendering configuration.</param>
    /// <param name="mode">Whether profile-owned settings replace or overlay current settings.</param>
    /// <returns>This options instance for fluent configuration.</returns>
    public static TOptions UseRenderingProfile<TOptions>(
        this TOptions options,
        OfficeRenderingProfile profile,
        OfficeRenderingProfileApplyMode mode = OfficeRenderingProfileApplyMode.Replace)
        where TOptions : OfficeImageExportOptions {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (profile == null) throw new ArgumentNullException(nameof(profile));
        if (mode != OfficeRenderingProfileApplyMode.Replace
            && mode != OfficeRenderingProfileApplyMode.Overlay) {
            throw new ArgumentOutOfRangeException(nameof(mode));
        }

        if (mode == OfficeRenderingProfileApplyMode.Replace) {
            options.Fonts = profile.FontsSnapshot.Clone();
            options.TextShapingProvider = profile.TextShapingProvider;
            options.TextShapingLanguage = profile.TextShapingLanguage;
            options.ImageCodec = profile.ImageCodec;
        } else {
            options.Fonts ??= new OfficeFontFaceCollection();
            options.Fonts.AddRangePreservingExisting(profile.FontsSnapshot);
            if (profile.TextShapingProvider != null) {
                options.TextShapingProvider = profile.TextShapingProvider;
            }
            if (profile.TextShapingLanguage != null) {
                options.TextShapingLanguage = profile.TextShapingLanguage;
            }
            if (profile.ImageCodec != null) {
                options.ImageCodec = profile.ImageCodec;
            }
        }

        options.Policy = profile.PolicySnapshot.Clone();
        return options;
    }
}
