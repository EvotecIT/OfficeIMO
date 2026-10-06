using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Gets or sets the minimum font size, in points, at which this run uses kerning.
        /// Null inherits the threshold; zero disables kerning. Values from zero to 1638
        /// are rounded to Word's half-point precision and are preserved in DOCX and native DOC.
        /// </summary>
        public double? KerningMinimumFontSizePoints {
            get => ScopedRunProperties?.Kern?.Val?.Value is uint halfPoints ? halfPoints / 2D : null;
            set {
                if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value) ||
                    value.Value < 0D || value.Value > 1638D)) {
                    throw new ArgumentOutOfRangeException(nameof(value));
                }

                RunProperties properties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = Hyperlink!;
                    properties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    properties = VerifyRunProperties();
                }

                if (value.HasValue) {
                    properties.Kern = new Kern {
                        Val = checked((uint)Math.Round(value.Value * 2D, MidpointRounding.AwayFromZero))
                    };
                } else {
                    properties.Kern?.Remove();
                }
            }
        }
    }
}
