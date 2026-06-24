using System.Globalization;
using System.Runtime.InteropServices;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Runtime diagnostics that can affect Excel import/export behavior.
    /// </summary>
    public sealed class ExcelRuntimePreflightReport {
        internal ExcelRuntimePreflightReport(
            string frameworkDescription,
            string osDescription,
            string currentCultureName,
            string currentUICultureName,
            bool globalizationInvariantMode,
            IReadOnlyList<string> warnings) {
            FrameworkDescription = frameworkDescription;
            OSDescription = osDescription;
            CurrentCultureName = currentCultureName;
            CurrentUICultureName = currentUICultureName;
            GlobalizationInvariantMode = globalizationInvariantMode;
            Warnings = warnings;
        }

        /// <summary>Gets the .NET framework description.</summary>
        public string FrameworkDescription { get; }

        /// <summary>Gets the operating system description.</summary>
        public string OSDescription { get; }

        /// <summary>Gets the current culture name.</summary>
        public string CurrentCultureName { get; }

        /// <summary>Gets the current UI culture name.</summary>
        public string CurrentUICultureName { get; }

        /// <summary>Gets a value indicating whether globalization-invariant mode appears to be enabled.</summary>
        public bool GlobalizationInvariantMode { get; }

        /// <summary>Gets runtime warnings relevant to Excel workflows.</summary>
        public IReadOnlyList<string> Warnings { get; }

        /// <summary>Gets a value indicating whether no runtime warnings were found.</summary>
        public bool IsClean => Warnings.Count == 0;
    }

    /// <summary>
    /// Inspects the current runtime for Excel workflow compatibility issues.
    /// </summary>
    public static class ExcelRuntimePreflight {
        /// <summary>
        /// Inspects the current process for culture/runtime settings that commonly affect spreadsheet workflows.
        /// </summary>
        public static ExcelRuntimePreflightReport InspectCurrent() {
            var warnings = new List<string>();
            bool invariantMode = IsGlobalizationInvariantMode();

            if (invariantMode) {
                warnings.Add("Globalization-invariant mode is enabled. Use explicit invariant formats or install ICU/globalization data before relying on culture-specific dates, numbers, or currency symbols.");
            }

            return new ExcelRuntimePreflightReport(
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                CultureInfo.CurrentCulture.Name,
                CultureInfo.CurrentUICulture.Name,
                invariantMode,
                warnings);
        }

        private static bool IsGlobalizationInvariantMode() {
            if (AppContext.TryGetSwitch("System.Globalization.Invariant", out bool invariantSwitch) && invariantSwitch) {
                return true;
            }

            string? environmentValue = Environment.GetEnvironmentVariable("DOTNET_SYSTEM_GLOBALIZATION_INVARIANT");
            if (string.Equals(environmentValue, "1", StringComparison.Ordinal)
                || string.Equals(environmentValue, "true", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return CultureInfo.CurrentCulture.Name.Length == 0
                && CultureInfo.CurrentUICulture.Name.Length == 0;
        }
    }
}
