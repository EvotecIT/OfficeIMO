using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Gets or sets the minimum font size, in points, at which this run uses kerning.
        /// Null inherits the threshold; zero disables kerning. Values from zero to 1638
        /// are rounded to Word's half-point precision and are preserved in DOCX and native DOC.
        /// </summary>
        public global::OfficeIMO.OfficeLength? KerningMinimumFontSize {
            get => ScopedRunProperties?.Kern?.Val?.Value is uint halfPoints ? global::OfficeIMO.OfficeLength.Emus((long)halfPoints * 6350) : null;
            set {
                if (value.HasValue && (value.Value < global::OfficeIMO.OfficeLength.Zero || value.Value > global::OfficeIMO.OfficeLength.Points(1638))) {
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
                        Val = checked((uint)value.Value.ToHalfPoints())
                    };
                } else {
                    properties.Kern?.Remove();
                }
            }
        }
    }
}
