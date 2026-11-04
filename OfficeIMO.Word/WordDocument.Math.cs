using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>Adds an editable Word equation authored with the reusable OfficeIMO math model.</summary>
        public WordEquation AddEquation(OfficeMathExpression expression) {
            ArgumentNullException.ThrowIfNull(expression);
            string omml = WordMathMarkup.ToOmml(expression);
            return AddParagraph().AddEquation(omml);
        }
    }
}
