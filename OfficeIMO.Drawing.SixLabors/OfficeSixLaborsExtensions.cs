using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Drawing.SixLabors;

/// <summary>Convenience activation methods for the optional managed font engine.</summary>
public static class OfficeSixLaborsExtensions {
    /// <summary>Enables the shared SixLabors font-program provider for subsequent face additions.</summary>
    public static OfficeFontFaceCollection UseSixLaborsFontPrograms(
        this OfficeFontFaceCollection fonts,
        OfficeSixLaborsFontProgramProvider? provider = null) {
        if (fonts == null) throw new ArgumentNullException(nameof(fonts));
        fonts.FontProgramProvider = provider ?? OfficeSixLaborsFontProgramProvider.Instance;
        return fonts;
    }

    /// <summary>Enables the shared SixLabors font-program provider on image or HTML render options.</summary>
    public static T UseSixLaborsFontPrograms<T>(
        this T options,
        OfficeSixLaborsFontProgramProvider? provider = null)
        where T : OfficeImageExportOptions {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Fonts ??= new OfficeFontFaceCollection();
        options.Fonts.UseSixLaborsFontPrograms(provider);
        return options;
    }
}
