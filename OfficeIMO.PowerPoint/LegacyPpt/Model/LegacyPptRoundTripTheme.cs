using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>
    /// Represents the DrawingML theme and color mapping stored in PowerPoint 2007+
    /// round-trip records inside a binary presentation.
    /// </summary>
    public sealed class LegacyPptRoundTripTheme {
        internal LegacyPptRoundTripTheme(string? themeXml,
            string? colorMappingXml, bool isOverride,
            string? name, string? colorSchemeName,
            IReadOnlyDictionary<PowerPointThemeColor, string> colors,
            string? majorLatinFont, string? minorLatinFont) {
            ThemeXml = themeXml;
            ColorMappingXml = colorMappingXml;
            IsOverride = isOverride;
            Name = name;
            ColorSchemeName = colorSchemeName;
            Colors = new ReadOnlyDictionary<PowerPointThemeColor, string>(
                colors.ToDictionary(pair => pair.Key, pair => pair.Value));
            MajorLatinFont = majorLatinFont;
            MinorLatinFont = minorLatinFont;
        }

        /// <summary>Gets the exact DrawingML theme or theme-override XML.</summary>
        public string? ThemeXml { get; }

        /// <summary>Gets the exact DrawingML color-mapping XML.</summary>
        public string? ColorMappingXml { get; }

        /// <summary>Gets whether the theme payload is a theme override.</summary>
        public bool IsOverride { get; }

        /// <summary>Gets the theme name, when the payload is a full theme.</summary>
        public string? Name { get; }

        /// <summary>Gets the DrawingML color-scheme name.</summary>
        public string? ColorSchemeName { get; }

        /// <summary>Gets directly resolvable DrawingML theme colors.</summary>
        public IReadOnlyDictionary<PowerPointThemeColor, string> Colors { get; }

        /// <summary>Gets the major Latin theme font.</summary>
        public string? MajorLatinFont { get; }

        /// <summary>Gets the minor Latin theme font.</summary>
        public string? MinorLatinFont { get; }
    }
}
