using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Identifies the native representation that backs a Word equation.</summary>
    public enum WordEquationRepresentation {
        /// <summary>Office Math Markup Language used by DOCX.</summary>
        Omml,
        /// <summary>Word EQ field used by legacy DOC and RTF.</summary>
        EquationField
    }

    /// <summary>
    /// Encapsulates an OMML equation or a legacy Word <c>EQ</c> field.
    /// </summary>
    public class WordEquation : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly DocumentFormat.OpenXml.Math.OfficeMath? _officeMath;
        private readonly DocumentFormat.OpenXml.Math.Paragraph? _mathParagraph;
        private readonly SimpleField? _simpleField;
        private readonly List<Run>? _runs;

        /// <summary>Initializes an equation backed by an Office Math element.</summary>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
            _document = document;
            _paragraph = paragraph;
            _officeMath = officeMath;
        }

        /// <summary>Initializes an equation backed by an Office Math paragraph.</summary>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            _document = document;
            _paragraph = paragraph;
            _mathParagraph = mathParagraph;
        }

        /// <summary>Initializes an equation backed by an Office Math element and its paragraph.</summary>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            _document = document;
            _paragraph = paragraph;
            _officeMath = officeMath;
            _mathParagraph = mathParagraph;
        }

        internal WordEquation(WordDocument document, Paragraph paragraph, SimpleField simpleField) {
            _document = document;
            _paragraph = paragraph;
            _simpleField = simpleField;
        }

        internal WordEquation(WordDocument document, Paragraph paragraph, List<Run> runs) {
            _document = document;
            _paragraph = paragraph;
            _runs = runs;
        }

        /// <summary>Gets the native equation representation.</summary>
        public WordEquationRepresentation Representation => MathElement == null
            ? WordEquationRepresentation.EquationField
            : WordEquationRepresentation.Omml;

        /// <summary>
        /// Gets or sets deterministic visible equation text. Setting OMML text replaces the structured
        /// expression with one math text run; setting an EQ field updates its cached display result.
        /// </summary>
        public string Text {
            get {
                OpenXmlElement? math = MathElement;
                if (math != null) return WordMath.GetText(math);
                return FieldParagraph?.Text ?? string.Empty;
            }
            set {
                OpenXmlElement? math = MathElement;
                if (math != null) {
                    WordMath.SetText(math, value);
                    return;
                }
                if (FieldParagraph != null) FieldParagraph.Text = value;
            }
        }

        /// <summary>Gets raw OMML when the equation is backed by DOCX math markup; otherwise <c>null</c>.</summary>
        public string? Omml => MathElement?.OuterXml;

        /// <summary>Gets the raw EQ field instruction when the equation is field-backed; otherwise <c>null</c>.</summary>
        public string? FieldInstruction => FieldParagraph?.Field?.Field;

        /// <summary>Projects the equation to LaTeX. Field-backed equations fall back to their cached display text.</summary>
        public string ToLatex() => MathElement != null ? WordMath.ToLatex(MathElement) : Text;

        /// <summary>Projects the equation to MathML. Field-backed equations use a safe <c>mtext</c> fallback.</summary>
        public string ToMathMl() => MathElement != null
            ? WordMath.ToMathMl(MathElement)
            : $"<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mtext>{EscapeXml(Text)}</mtext></math>";

        /// <summary>Projects OMML to a legacy Word EQ instruction, or returns the existing EQ instruction.</summary>
        public string ToEquationFieldInstruction() => MathElement != null
            ? WordMath.ToEquationFieldInstruction(MathElement)
            : FieldInstruction ?? " EQ ";

        /// <summary>Removes the equation and its backing markup from the document.</summary>
        public void Remove() {
            _officeMath?.Remove();
            _mathParagraph?.Remove();
            _simpleField?.Remove();
            if (_runs != null) {
                foreach (Run run in _runs.ToList()) run.Remove();
            }
        }

        internal OpenXmlElement? MathElement => (OpenXmlElement?)_officeMath ?? _mathParagraph;

        private WordParagraph? FieldParagraph => _simpleField != null
            ? new WordParagraph(_document, _paragraph, _simpleField)
            : _runs != null ? new WordParagraph(_document, _paragraph, _runs) : null;

        private static string EscapeXml(string value) => value
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }
}
