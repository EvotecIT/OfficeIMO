using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>Adds an editable Word equation authored with the reusable OfficeIMO math model.</summary>
        public WordParagraph AddEquation(OfficeMathExpression expression) {
            if (expression == null) throw new ArgumentNullException(nameof(expression));
            return AddParagraph().AddEquation(expression);
        }
    }
}
