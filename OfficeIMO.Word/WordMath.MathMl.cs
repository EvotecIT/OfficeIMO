using DocumentFormat.OpenXml;
using System.Text;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Word {
    internal static partial class WordMath {
        internal static string ToMathMl(OpenXmlElement element) {
            var builder = new StringBuilder();
            builder.Append("<math xmlns=\"http://www.w3.org/1998/Math/MathML\">");
            AppendMathMl(builder, element);
            builder.Append("</math>");
            return builder.ToString();
        }
        private static void AppendMathMl(StringBuilder builder, OpenXmlElement element) {
            if (element is M.Text text) {
                builder.Append("<mtext>");
                builder.Append(EscapeXml(text.Text));
                builder.Append("</mtext>");
                return;
            }

            switch (element.LocalName) {
                case "f":
                    AppendMathMlTwoChildElement(builder, "mfrac", element, "num", "den");
                    return;
                case "sSup":
                    AppendMathMlTwoChildElement(builder, "msup", element, "e", "sup");
                    return;
                case "sSub":
                    AppendMathMlTwoChildElement(builder, "msub", element, "e", "sub");
                    return;
                case "sSubSup":
                    AppendMathMlThreeChildElement(builder, "msubsup", element, "e", "sub", "sup");
                    return;
                case "sPre":
                    builder.Append("<mmultiscripts>");
                    AppendMathMlChild(builder, element, "e");
                    builder.Append("<mprescripts/>");
                    AppendMathMlChildOrNone(builder, element, "sub");
                    AppendMathMlChildOrNone(builder, element, "sup");
                    builder.Append("</mmultiscripts>");
                    return;
                case "rad":
                    OpenXmlElement? degree = FindFirstChild(element, "deg");
                    if (degree == null || GetText(degree).Length == 0) {
                        builder.Append("<msqrt>");
                        AppendMathMlChild(builder, element, "e");
                        builder.Append("</msqrt>");
                    } else {
                        builder.Append("<mroot>");
                        AppendMathMlChild(builder, element, "e");
                        AppendMathMl(builder, degree);
                        builder.Append("</mroot>");
                    }
                    return;
                case "nary":
                case "int":
                    AppendNaryMathMl(builder, element);
                    return;
                case "func":
                    builder.Append("<mrow><mi>");
                    builder.Append(EscapeXml(ReadChildText(element, "fName")));
                    builder.Append("</mi><mo>(</mo>");
                    AppendMathMlChild(builder, element, "e");
                    builder.Append("<mo>)</mo></mrow>");
                    return;
                case "acc":
                    builder.Append("<mover accent=\"true\">");
                    AppendMathMlChild(builder, element, "e");
                    AppendMathMlOperator(builder, ReadCharacter(element, "chr").Value);
                    builder.Append("</mover>");
                    return;
                case "bar":
                    builder.Append("<mover accent=\"true\">");
                    AppendMathMlChild(builder, element, "e");
                    AppendMathMlOperator(builder, "¯");
                    builder.Append("</mover>");
                    return;
                case "d":
                    AppendDelimiterMathMl(builder, element);
                    return;
                case "groupChr":
                    builder.Append("<mover>");
                    AppendMathMlChild(builder, element, "e");
                    AppendMathMlOperator(builder, ReadCharacter(element, "chr").Value);
                    builder.Append("</mover>");
                    return;
                case "m":
                    AppendMatrixMathMl(builder, element);
                    return;
                case "eqArr":
                    builder.Append("<mtable>");
                    foreach (OpenXmlElement expression in FindChildren(element, "e")) {
                        builder.Append("<mtr><mtd>");
                        AppendMathMl(builder, expression);
                        builder.Append("</mtd></mtr>");
                    }
                    builder.Append("</mtable>");
                    return;
                case "limLow":
                    AppendMathMlTwoChildElement(builder, "munder", element, "e", "lim");
                    return;
                case "limUpp":
                    AppendMathMlTwoChildElement(builder, "mover", element, "e", "lim");
                    return;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendMathMl(builder, child);
            }
        }

        private static void AppendMathMlTwoChildElement(StringBuilder builder, string tag, OpenXmlElement element, string first, string second) {
            builder.Append('<').Append(tag).Append('>');
            AppendMathMlChild(builder, element, first);
            AppendMathMlChild(builder, element, second);
            builder.Append("</").Append(tag).Append('>');
        }

        private static void AppendMathMlThreeChildElement(StringBuilder builder, string tag, OpenXmlElement element, string first, string second, string third) {
            builder.Append('<').Append(tag).Append('>');
            AppendMathMlChild(builder, element, first);
            AppendMathMlChild(builder, element, second);
            AppendMathMlChild(builder, element, third);
            builder.Append("</").Append(tag).Append('>');
        }

        private static void AppendMathMlChild(StringBuilder builder, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            if (child != null) AppendMathMl(builder, child);
        }

        private static void AppendMathMlChildOrNone(StringBuilder builder, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            if (child == null || GetText(child).Length == 0) builder.Append("<none/>");
            else AppendMathMl(builder, child);
        }

        private static void AppendNaryMathMl(StringBuilder builder, OpenXmlElement element) {
            builder.Append("<mrow><munderover>");
            AppendMathMlOperator(builder, ReadNaryOperatorText(element) switch {
                "sum" => "∑",
                "prod" => "∏",
                "int" => "∫",
                string other => other
            });
            AppendMathMlChildOrNone(builder, element, "sub");
            AppendMathMlChildOrNone(builder, element, "sup");
            builder.Append("</munderover>");
            AppendMathMlChild(builder, element, "e");
            builder.Append("</mrow>");
        }

        private static void AppendDelimiterMathMl(StringBuilder builder, OpenXmlElement element) {
            MathCharacter begin = ReadCharacter(element, "begChr");
            MathCharacter end = ReadCharacter(element, "endChr");
            builder.Append("<mrow>");
            string beginValue = begin.Present ? begin.Value : "(";
            if (beginValue.Length > 0) AppendMathMlOperator(builder, beginValue, fence: true);
            bool first = true;
            foreach (OpenXmlElement expression in FindChildren(element, "e")) {
                if (!first) AppendMathMlOperator(builder, ",");
                AppendMathMl(builder, expression);
                first = false;
            }
            string endValue = end.Present ? end.Value : ")";
            if (endValue.Length > 0) AppendMathMlOperator(builder, endValue, fence: true);
            builder.Append("</mrow>");
        }

        private static void AppendMatrixMathMl(StringBuilder builder, OpenXmlElement element) {
            builder.Append("<mtable>");
            foreach (OpenXmlElement row in FindChildren(element, "mr")) {
                builder.Append("<mtr>");
                foreach (OpenXmlElement cell in FindChildren(row, "e")) {
                    builder.Append("<mtd>");
                    AppendMathMl(builder, cell);
                    builder.Append("</mtd>");
                }
                builder.Append("</mtr>");
            }
            builder.Append("</mtable>");
        }

        private static void AppendMathMlOperator(StringBuilder builder, string value, bool fence = false) {
            builder.Append(fence ? "<mo fence=\"true\">" : "<mo>");
            builder.Append(EscapeXml(value));
            builder.Append("</mo>");
        }

        private static string EscapeXml(string? value) => (value ?? string.Empty)
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");

    }
}
