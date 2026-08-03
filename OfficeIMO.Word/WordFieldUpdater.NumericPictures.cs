using System;
using System.Globalization;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        /// <summary>
        /// Validates every selectable section of a numeric picture before a caller accepts the
        /// formatting profile for values that are not yet known.
        /// </summary>
        internal static bool TryValidateNumericPictureProfile(string? numericPicture, out string? diagnostic) {
            diagnostic = null;
            if (string.IsNullOrWhiteSpace(numericPicture)) return true;

            string format = TrimFormulaFormatQuotes(numericPicture!);
            if (format.Length == 0) {
                diagnostic = "Formula numeric picture switch is empty.";
                return false;
            }

            string[] sectionTexts = SplitNumericPictureSections(format);
            if (sectionTexts.Length > 3) {
                diagnostic = $"Formula numeric picture switch '{format}' contains more than three sections.";
                return false;
            }

            foreach (string sectionText in sectionTexts) {
                if (!TryParseNumericPictureSection(sectionText, out NumericPictureSection section, out diagnostic)) {
                    return false;
                }
                if (section.Format.Length == 0) {
                    diagnostic = $"Formula numeric picture switch '{format}' contains an empty section.";
                    return false;
                }
                if (!TryNormalizeNumericPictureFillSyntax(section.Format, out string normalizedFormat, out diagnostic)) {
                    return false;
                }
                if (normalizedFormat.Length == 0) {
                    diagnostic = $"Formula numeric picture switch '{format}' contains a section with only layout-dependent fill formatting syntax.";
                    return false;
                }
                if (!TryValidateNumericPicture(normalizedFormat, out diagnostic)) {
                    return false;
                }

                try {
                    _ = 0m.ToString(normalizedFormat, CultureInfo.InvariantCulture);
                } catch (FormatException) {
                    diagnostic = $"Formula numeric picture section '{sectionText}' is not supported by the bounded OfficeIMO formatter.";
                    return false;
                }
            }

            return true;
        }
    }
}
