namespace OfficeIMO.Excel {
    /// <summary>
    /// Compatibility profile used when selecting Excel table styles.
    /// </summary>
    public enum ExcelTableStyleCompatibilityProfile {
        /// <summary>
        /// Accept any built-in Excel table style.
        /// </summary>
        Desktop,

        /// <summary>
        /// Prefer table styles that are broadly stable across desktop, web, and spreadsheet viewers.
        /// </summary>
        CrossHost
    }

    /// <summary>
    /// Describes how a table style fits a compatibility profile.
    /// </summary>
    public sealed class ExcelTableStyleCompatibilityInfo {
        internal ExcelTableStyleCompatibilityInfo(
            ExcelTableStyle style,
            string name,
            ExcelTableStyleCompatibilityProfile profile,
            bool isBuiltIn,
            bool isRecommended,
            string? warning) {
            Style = style;
            Name = name;
            Profile = profile;
            IsBuiltIn = isBuiltIn;
            IsRecommended = isRecommended;
            Warning = warning;
        }

        /// <summary>Gets the table style enum value.</summary>
        public ExcelTableStyle Style { get; }

        /// <summary>Gets the workbook style name written to table XML.</summary>
        public string Name { get; }

        /// <summary>Gets the profile used for the compatibility check.</summary>
        public ExcelTableStyleCompatibilityProfile Profile { get; }

        /// <summary>Gets a value indicating whether the style is a known built-in Excel table style.</summary>
        public bool IsBuiltIn { get; }

        /// <summary>Gets a value indicating whether the style is recommended for the selected profile.</summary>
        public bool IsRecommended { get; }

        /// <summary>Gets a warning when the style is valid but not recommended for the selected profile.</summary>
        public string? Warning { get; }
    }

    /// <summary>
    /// Provides built-in Excel table style names and compatibility recommendations.
    /// </summary>
    public static class ExcelTableStyleCatalog {
        private static readonly HashSet<ExcelTableStyle> CrossHostRecommendedStyles = new HashSet<ExcelTableStyle> {
            ExcelTableStyle.TableStyleLight1,
            ExcelTableStyle.TableStyleLight2,
            ExcelTableStyle.TableStyleLight9,
            ExcelTableStyle.TableStyleLight11,
            ExcelTableStyle.TableStyleMedium2,
            ExcelTableStyle.TableStyleMedium4,
            ExcelTableStyle.TableStyleMedium9,
            ExcelTableStyle.TableStyleMedium15,
            ExcelTableStyle.TableStyleMedium21
        };

        /// <summary>
        /// Returns all built-in Excel table style names.
        /// </summary>
        public static IReadOnlyList<string> GetNames() {
            return Enum.GetNames(typeof(ExcelTableStyle));
        }

        /// <summary>
        /// Returns table style names recommended for the given compatibility profile.
        /// </summary>
        public static IReadOnlyList<string> GetRecommendedNames(ExcelTableStyleCompatibilityProfile profile = ExcelTableStyleCompatibilityProfile.CrossHost) {
            if (profile == ExcelTableStyleCompatibilityProfile.Desktop) {
                return GetNames();
            }

            return CrossHostRecommendedStyles
                .OrderBy(style => style.ToString(), StringComparer.Ordinal)
                .Select(style => style.ToString())
                .ToArray();
        }

        /// <summary>
        /// Tries to parse a built-in table style name.
        /// </summary>
        public static bool TryParse(string? name, out ExcelTableStyle style) {
            return Enum.TryParse(name, ignoreCase: true, out style)
                && Enum.IsDefined(typeof(ExcelTableStyle), style);
        }

        /// <summary>
        /// Gets compatibility information for a built-in table style.
        /// </summary>
        public static ExcelTableStyleCompatibilityInfo Analyze(
            ExcelTableStyle style,
            ExcelTableStyleCompatibilityProfile profile = ExcelTableStyleCompatibilityProfile.CrossHost) {
            bool recommended = profile == ExcelTableStyleCompatibilityProfile.Desktop
                || CrossHostRecommendedStyles.Contains(style);
            string? warning = recommended
                ? null
                : "The table style is valid, but a simpler light or medium style is recommended for cross-host workbook compatibility.";

            return new ExcelTableStyleCompatibilityInfo(
                style,
                style.ToString(),
                profile,
                isBuiltIn: true,
                isRecommended: recommended,
                warning);
        }

        /// <summary>
        /// Gets compatibility information for a table style name.
        /// </summary>
        public static ExcelTableStyleCompatibilityInfo Analyze(
            string name,
            ExcelTableStyleCompatibilityProfile profile = ExcelTableStyleCompatibilityProfile.CrossHost) {
            if (!TryParse(name, out ExcelTableStyle style)) {
                return new ExcelTableStyleCompatibilityInfo(
                    default,
                    name,
                    profile,
                    isBuiltIn: false,
                    isRecommended: false,
                    "The table style name is not one of the built-in Excel table styles.");
            }

            return Analyze(style, profile);
        }
    }
}
