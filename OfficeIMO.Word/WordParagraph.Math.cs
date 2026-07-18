using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Adds an inline editable Word equation authored with the reusable OfficeIMO math model.</summary>
        public WordParagraph AddEquation(OfficeMathExpression expression) {
            if (expression == null) throw new ArgumentNullException(nameof(expression));
            return AddEquation(WordMathMarkup.ToOmml(expression));
        }
    }
}
