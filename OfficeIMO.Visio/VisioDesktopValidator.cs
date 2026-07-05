using System;
using System.IO;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Compatibility surface for older callers that probed desktop validation availability.
    /// OfficeIMO.Visio rendering and validation remain dependency-free and do not automate desktop applications.
    /// </summary>
    public static class VisioDesktopValidator {
        private const string DependencyFreeMessage = "Desktop application automation is not part of OfficeIMO.Visio dependency-free validation.";

        /// <summary>
        /// Gets whether optional desktop validation is available through the OfficeIMO.Visio package.
        /// </summary>
        public static bool IsAvailable() => false;

        /// <summary>
        /// Validates that the target path exists, then reports that desktop validation is unavailable.
        /// </summary>
        /// <param name="vsdxPath">Path to the VSDX package.</param>
        /// <returns>Desktop validation result.</returns>
        public static VisioDesktopValidationResult Validate(string vsdxPath) {
            return Validate(vsdxPath, null);
        }

        /// <summary>
        /// Validates that the target path exists, then reports that desktop validation is unavailable.
        /// </summary>
        /// <param name="vsdxPath">Path to the VSDX package.</param>
        /// <param name="options">Ignored compatibility options.</param>
        /// <returns>Desktop validation result.</returns>
        public static VisioDesktopValidationResult Validate(string vsdxPath, VisioDesktopValidationOptions? options) {
            if (string.IsNullOrWhiteSpace(vsdxPath)) {
                throw new ArgumentException("VSDX path cannot be null or whitespace.", nameof(vsdxPath));
            }

            string fullPath = Path.GetFullPath(vsdxPath);
            if (!File.Exists(fullPath)) {
                throw new FileNotFoundException("VSDX file was not found.", fullPath);
            }

            return new VisioDesktopValidationResult(
                isAvailable: false,
                isValid: false,
                version: null,
                issues: new[] { DependencyFreeMessage });
        }
    }
}
