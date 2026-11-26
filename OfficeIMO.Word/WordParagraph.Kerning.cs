using static OfficeIMO.Word.OpenXml.WordXmlEdits;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Gets or sets the minimum font size, in points, at which this run uses kerning.
        /// Null inherits the threshold; zero disables kerning. Values from zero to 1638
        /// are rounded to Word's half-point precision and are preserved in DOCX and native DOC.
        /// </summary>
        public global::OfficeIMO.OfficeLength? KerningMinimumFontSize {
            get => uint.TryParse(ChildValue(RunPropertiesNode, "kern"), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out uint halfPoints)
                ? global::OfficeIMO.OfficeLength.Emus((long)halfPoints * 6350) : null;
            set {
                if (value.HasValue && (value.Value < global::OfficeIMO.OfficeLength.Zero || value.Value > global::OfficeIMO.OfficeLength.Points(1638))) {
                    throw new ArgumentOutOfRangeException(nameof(value));
                }
                SetRunValue("kern", value.HasValue
                    ? checked((uint)value.Value.ToHalfPoints()).ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : null);
            }
        }
    }
}
