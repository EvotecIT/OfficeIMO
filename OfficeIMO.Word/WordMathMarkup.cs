using OfficeIMO.Drawing;
using System.Xml.Linq;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Word {
    /// <summary>Converts between Word's native OMML and the reusable OfficeIMO math expression tree.</summary>
    public static class WordMathMarkup {
        /// <summary>Serializes a shared expression as inline <c>m:oMath</c> or display <c>m:oMathPara</c> markup.</summary>
        public static string ToOmml(OfficeMathExpression expression, bool display = false) => WordMath.ToOmml(expression, display);

        /// <summary>Parses an inline or display OMML fragment into the reusable math expression tree.</summary>
        public static OfficeMathExpression FromOmml(string omml) => FromOmml(omml, OfficeMathMarkup.DefaultMaximumParseDepth);

        /// <summary>Parses OMML with a hard nesting-depth limit.</summary>
        public static OfficeMathExpression FromOmml(string omml, int maximumDepth) {
            if (string.IsNullOrWhiteSpace(omml)) throw new ArgumentException("OMML cannot be empty.", nameof(omml));
            if (maximumDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumDepth));
            OfficeMathMarkup.ValidateXmlDepth(omml, maximumDepth);
            XElement root = XElement.Parse(omml, LoadOptions.PreserveWhitespace);
            if (root.Name.LocalName == "oMath") return WordMath.ToExpression(new M.OfficeMath(omml));
            if (root.Name.LocalName == "oMathPara") return WordMath.ToExpression(new M.Paragraph(omml));
            throw new FormatException("The OMML root must be m:oMath or m:oMathPara.");
        }
    }
}
