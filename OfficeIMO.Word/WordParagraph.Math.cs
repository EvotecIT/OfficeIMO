using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Adds an inline editable Word equation authored with the reusable OfficeIMO math model.</summary>
        public WordEquation AddEquation(OfficeMathExpression expression) {
            ArgumentNullException.ThrowIfNull(expression);
            return AddEquation(WordMathMarkup.ToOmml(expression));
        }
    }
}
