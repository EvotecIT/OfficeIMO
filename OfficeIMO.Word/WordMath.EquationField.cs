using DocumentFormat.OpenXml;
using System.Text;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Word {
    internal static partial class WordMath {
        internal static string ToEquationFieldInstruction(OpenXmlElement element) {
            var builder = new StringBuilder(" EQ ");
            AppendEquationField(builder, element);
            builder.Append(" \\* MERGEFORMAT ");
            return builder.ToString();
        }
        private static void AppendEquationField(StringBuilder builder, OpenXmlElement element) {
            if (element is M.Text text) {
                builder.Append(EscapeEquationFieldLiteral(text.Text));
                return;
            }

            switch (element.LocalName) {
                case "f":
                    builder.Append("\\f(");
                    AppendEquationFieldChild(builder, element, "num");
                    builder.Append(',');
                    AppendEquationFieldChild(builder, element, "den");
                    builder.Append(')');
                    return;
                case "sSup":
                    AppendEquationFieldChild(builder, element, "e");
                    AppendEquationFieldScript(builder, "\\up8", element, "sup");
                    return;
                case "sSub":
                    AppendEquationFieldChild(builder, element, "e");
                    AppendEquationFieldScript(builder, "\\do8", element, "sub");
                    return;
                case "sSubSup":
                    AppendEquationFieldChild(builder, element, "e");
                    AppendEquationFieldScript(builder, "\\do8", element, "sub");
                    AppendEquationFieldScript(builder, "\\up8", element, "sup");
                    return;
                case "sPre":
                    AppendEquationFieldScript(builder, "\\up8", element, "sup");
                    AppendEquationFieldScript(builder, "\\do8", element, "sub");
                    AppendEquationFieldChild(builder, element, "e");
                    return;
                case "rad":
                    builder.Append("\\r(");
                    AppendEquationFieldChild(builder, element, "deg");
                    builder.Append(',');
                    AppendEquationFieldChild(builder, element, "e");
                    builder.Append(')');
                    return;
                case "nary":
                case "int":
                    AppendNaryEquationField(builder, element);
                    return;
                case "func":
                    AppendEquationFieldChild(builder, element, "fName");
                    builder.Append("\\(");
                    AppendEquationFieldChild(builder, element, "e");
                    builder.Append(')');
                    return;
                case "acc":
                    builder.Append("\\o(");
                    AppendEquationFieldChild(builder, element, "e");
                    builder.Append(',');
                    builder.Append(EscapeEquationFieldLiteral(ReadCharacter(element, "chr").Value));
                    builder.Append(')');
                    return;
                case "bar":
                    builder.Append("\\x");
                    builder.Append(string.Equals(ReadCharacter(element, "pos").Value, "bot", StringComparison.OrdinalIgnoreCase)
                        ? "\\bo("
                        : "\\to(");
                    AppendEquationFieldChild(builder, element, "e");
                    builder.Append(')');
                    return;
                case "d":
                    AppendDelimiterEquationField(builder, element);
                    return;
                case "groupChr":
                    builder.Append(EscapeEquationFieldLiteral(GetText(element)));
                    return;
                case "m":
                    AppendMatrixEquationField(builder, element);
                    return;
                case "eqArr":
                    AppendEquationArrayField(builder, element);
                    return;
                case "limLow":
                    AppendEquationFieldChild(builder, element, "e");
                    AppendEquationFieldScript(builder, "\\do8", element, "lim");
                    return;
                case "limUpp":
                    AppendEquationFieldChild(builder, element, "e");
                    AppendEquationFieldScript(builder, "\\up8", element, "lim");
                    return;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendEquationField(builder, child);
            }
        }

        private static void AppendEquationFieldChild(StringBuilder builder, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            if (child != null) AppendEquationField(builder, child);
        }

        private static void AppendEquationFieldScript(StringBuilder builder, string switchName, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            if (child == null || GetText(child).Length == 0) return;
            builder.Append("\\s");
            builder.Append(switchName);
            builder.Append('(');
            AppendEquationField(builder, child);
            builder.Append(')');
        }

        private static void AppendNaryEquationField(StringBuilder builder, OpenXmlElement element) {
            string operatorText = ReadNaryOperatorText(element);
            builder.Append("\\i");
            if (operatorText == "sum") builder.Append("\\su");
            else if (operatorText == "prod") builder.Append("\\pr");
            else if (operatorText != "int") builder.Append("\\fc").Append(EscapeEquationFieldDelimiter(operatorText));
            builder.Append('(');
            AppendEquationFieldChild(builder, element, "sub");
            builder.Append(',');
            AppendEquationFieldChild(builder, element, "sup");
            builder.Append(',');
            AppendEquationFieldChild(builder, element, "e");
            builder.Append(')');
        }

        private static void AppendDelimiterEquationField(StringBuilder builder, OpenXmlElement element) {
            MathCharacter begin = ReadCharacter(element, "begChr");
            MathCharacter end = ReadCharacter(element, "endChr");
            string beginValue = begin.Present ? begin.Value : "(";
            string endValue = end.Present ? end.Value : ")";
            if (beginValue.Length == 0 || endValue.Length == 0) {
                builder.Append(EscapeEquationFieldLiteral(beginValue));
                AppendJoinedEquationFieldChildren(builder, element, "e", "\\,");
                builder.Append(EscapeEquationFieldLiteral(endValue));
                return;
            }
            builder.Append("\\b\\lc");
            builder.Append(EscapeEquationFieldDelimiter(beginValue));
            builder.Append("\\rc");
            builder.Append(EscapeEquationFieldDelimiter(endValue));
            builder.Append('(');
            AppendJoinedEquationFieldChildren(builder, element, "e", "\\,");
            builder.Append(')');
        }

        private static void AppendMatrixEquationField(StringBuilder builder, OpenXmlElement element) {
            List<OpenXmlElement> rows = FindChildren(element, "mr").ToList();
            int columns = rows.Count == 0 ? 1 : Math.Max(1, FindChildren(rows[0], "e").Count());
            builder.Append("\\a\\co").Append(columns).Append('(');
            bool first = true;
            foreach (OpenXmlElement row in rows) {
                foreach (OpenXmlElement cell in FindChildren(row, "e")) {
                    if (!first) builder.Append(',');
                    AppendEquationField(builder, cell);
                    first = false;
                }
            }
            builder.Append(')');
        }

        private static void AppendEquationArrayField(StringBuilder builder, OpenXmlElement element) {
            builder.Append("\\a\\co1(");
            AppendJoinedEquationFieldChildren(builder, element, "e", ",");
            builder.Append(')');
        }

        private static void AppendJoinedEquationFieldChildren(StringBuilder builder, OpenXmlElement element, string localName, string separator) {
            bool first = true;
            foreach (OpenXmlElement child in FindChildren(element, localName)) {
                if (!first) builder.Append(separator);
                AppendEquationField(builder, child);
                first = false;
            }
        }

        private static string EscapeEquationFieldLiteral(string? value) => (value ?? string.Empty)
            .Replace("\\", "\\\\")
            .Replace("(", "\\(")
            .Replace(")", "\\)")
            .Replace(",", "\\,");

        private static string EscapeEquationFieldDelimiter(string value) => value == "\\" ? "\\\\" : "\\" + value;

    }
}
