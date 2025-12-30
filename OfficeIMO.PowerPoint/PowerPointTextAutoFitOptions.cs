using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Options for text auto-fit behavior.
    /// </summary>
    public readonly struct PowerPointTextAutoFitOptions {
        /// <summary>
        ///     Creates auto-fit options using percentage values (0-100).
        /// </summary>
        public PowerPointTextAutoFitOptions(double? fontScalePercent = null, double? lineSpaceReductionPercent = null) {
            ValidatePercent(fontScalePercent, nameof(fontScalePercent));
            ValidatePercent(lineSpaceReductionPercent, nameof(lineSpaceReductionPercent));
            FontScalePercent = fontScalePercent;
            LineSpaceReductionPercent = lineSpaceReductionPercent;
        }

        /// <summary>
        ///     Target font scale in percent (0-100).
        /// </summary>
        public double? FontScalePercent { get; }

        /// <summary>
        ///     Line spacing reduction in percent (0-100).
        /// </summary>
        public double? LineSpaceReductionPercent { get; }

        internal int? FontScaleValue => ToOpenXmlPercent(FontScalePercent);
        internal int? LineSpaceReductionValue => ToOpenXmlPercent(LineSpaceReductionPercent);

        internal static PowerPointTextAutoFitOptions FromOpenXmlValues(int? fontScale, int? lineSpaceReduction) {
            double? fontScalePercent = fontScale.HasValue ? fontScale.Value / 1000d : null;
            double? lineSpaceReductionPercent = lineSpaceReduction.HasValue ? lineSpaceReduction.Value / 1000d : null;
            return new PowerPointTextAutoFitOptions(fontScalePercent, lineSpaceReductionPercent);
        }

        private static void ValidatePercent(double? percent, string paramName) {
            if (percent == null) {
                return;
            }
            if (percent < 0 || percent > 100) {
                throw new ArgumentOutOfRangeException(paramName, "Percent must be between 0 and 100.");
            }
        }

        private static int? ToOpenXmlPercent(double? percent) {
            if (percent == null) {
                return null;
            }
            return (int)Math.Round(percent.Value * 1000d);
        }

        /// <summary>
        ///     Returns a display-friendly string.
        /// </summary>
        public override string ToString() {
            string fontScale = FontScalePercent?.ToString("0.###") ?? "?";
            string lineSpace = LineSpaceReductionPercent?.ToString("0.###") ?? "?";
            return $"{fontScale}% / {lineSpace}%";
        }
    }
}
