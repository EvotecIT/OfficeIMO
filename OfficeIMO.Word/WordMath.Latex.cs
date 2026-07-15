using DocumentFormat.OpenXml;
using System.Text;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Word {
    internal static partial class WordMath {
        internal static string ToLatex(OpenXmlElement element) {
            var builder = new StringBuilder();
            AppendLatex(builder, element);
            return builder.ToString();
        }
        private static void AppendLatex(StringBuilder builder, OpenXmlElement element) {
            if (element is M.Text text) {
                builder.Append(EscapeLatex(text.Text));
                return;
            }

            switch (element.LocalName) {
                case "f":
                    AppendLatexCommandWithTwoChildren(builder, "\\frac", element, "num", "den");
                    return;
                case "sSup":
                    AppendLatexChild(builder, element, "e");
                    AppendLatexScript(builder, '^', element, "sup");
                    return;
                case "sSub":
                    AppendLatexChild(builder, element, "e");
                    AppendLatexScript(builder, '_', element, "sub");
                    return;
                case "sSubSup":
                    AppendLatexChild(builder, element, "e");
                    AppendLatexScript(builder, '_', element, "sub");
                    AppendLatexScript(builder, '^', element, "sup");
                    return;
                case "sPre":
                    builder.Append("{} ");
                    AppendLatexScript(builder, '^', element, "sup");
                    AppendLatexScript(builder, '_', element, "sub");
                    AppendLatexChild(builder, element, "e");
                    return;
                case "rad":
                    OpenXmlElement? degree = FindFirstChild(element, "deg");
                    builder.Append("\\sqrt");
                    if (degree != null && GetText(degree).Length > 0) {
                        builder.Append('[');
                        AppendLatex(builder, degree);
                        builder.Append(']');
                    }
                    builder.Append('{');
                    AppendLatexChild(builder, element, "e");
                    builder.Append('}');
                    return;
                case "nary":
                case "int":
                    AppendNaryLatex(builder, element);
                    return;
                case "func":
                    builder.Append("\\operatorname{");
                    AppendLatexChild(builder, element, "fName");
                    builder.Append("}\\left(");
                    AppendLatexChild(builder, element, "e");
                    builder.Append("\\right)");
                    return;
                case "acc":
                    AppendAccentLatex(builder, element);
                    return;
                case "bar":
                    builder.Append(ReadCharacter(element, "pos").Value == "bot" ? "\\underline{" : "\\overline{");
                    AppendLatexChild(builder, element, "e");
                    builder.Append('}');
                    return;
                case "d":
                    AppendDelimiterLatex(builder, element);
                    return;
                case "groupChr":
                    AppendGroupCharacterLatex(builder, element);
                    return;
                case "m":
                    AppendMatrixLatex(builder, element);
                    return;
                case "eqArr":
                    AppendEquationArrayLatex(builder, element);
                    return;
                case "limLow":
                    AppendLatexChild(builder, element, "e");
                    AppendLatexScript(builder, '_', element, "lim");
                    return;
                case "limUpp":
                    AppendLatexChild(builder, element, "e");
                    AppendLatexScript(builder, '^', element, "lim");
                    return;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendLatex(builder, child);
            }
        }

        private static void AppendLatexCommandWithTwoChildren(StringBuilder builder, string command, OpenXmlElement element, string first, string second) {
            builder.Append(command);
            builder.Append('{');
            AppendLatexChild(builder, element, first);
            builder.Append("}{");
            AppendLatexChild(builder, element, second);
            builder.Append('}');
        }

        private static void AppendLatexChild(StringBuilder builder, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            if (child != null) AppendLatex(builder, child);
        }

        private static void AppendLatexScript(StringBuilder builder, char marker, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            if (child == null || GetText(child).Length == 0) return;
            builder.Append(marker);
            builder.Append('{');
            AppendLatex(builder, child);
            builder.Append('}');
        }

        private static void AppendNaryLatex(StringBuilder builder, OpenXmlElement element) {
            string operatorText = ReadNaryOperatorText(element);
            builder.Append(operatorText switch {
                "sum" => "\\sum",
                "prod" => "\\prod",
                "int" => "\\int",
                _ => "\\operatorname{" + EscapeLatex(operatorText) + "}"
            });
            AppendLatexScript(builder, '_', element, "sub");
            AppendLatexScript(builder, '^', element, "sup");
            OpenXmlElement? expression = FindFirstChild(element, "e");
            if (expression != null && GetText(expression).Length > 0) {
                builder.Append(" ");
                AppendLatex(builder, expression);
            }
        }

        private static void AppendAccentLatex(StringBuilder builder, OpenXmlElement element) {
            string accent = ReadCharacter(element, "chr").Value;
            string command = accent switch {
                "^" => "\\hat",
                "\u0302" => "\\hat",
                "~" => "\\tilde",
                "\u0303" => "\\tilde",
                "." => "\\dot",
                "\u0307" => "\\dot",
                "\u00a8" => "\\ddot",
                "\u0308" => "\\ddot",
                _ => "\\overset{" + EscapeLatex(accent) + "}"
            };
            builder.Append(command);
            builder.Append('{');
            AppendLatexChild(builder, element, "e");
            builder.Append('}');
        }

        private static void AppendDelimiterLatex(StringBuilder builder, OpenXmlElement element) {
            MathCharacter begin = ReadCharacter(element, "begChr");
            MathCharacter end = ReadCharacter(element, "endChr");
            builder.Append("\\left");
            builder.Append(ToLatexDelimiter(begin.Present ? begin.Value : "("));
            bool first = true;
            foreach (OpenXmlElement expression in FindChildren(element, "e")) {
                if (!first) builder.Append(',');
                AppendLatex(builder, expression);
                first = false;
            }
            builder.Append("\\right");
            builder.Append(ToLatexDelimiter(end.Present ? end.Value : ")"));
        }

        private static void AppendGroupCharacterLatex(StringBuilder builder, OpenXmlElement element) {
            string character = ReadCharacter(element, "chr").Value;
            string command = character switch {
                "\u23de" => "\\overbrace",
                "\u23df" => "\\underbrace",
                "\u23b4" => "\\overbracket",
                "\u23b5" => "\\underbracket",
                _ => "\\overbrace"
            };
            builder.Append(command);
            builder.Append('{');
            AppendLatexChild(builder, element, "e");
            builder.Append('}');
        }

        private static void AppendMatrixLatex(StringBuilder builder, OpenXmlElement element) {
            builder.Append("\\begin{matrix}");
            bool firstRow = true;
            foreach (OpenXmlElement row in FindChildren(element, "mr")) {
                if (!firstRow) builder.Append(" \\\\ ");
                bool firstCell = true;
                foreach (OpenXmlElement cell in FindChildren(row, "e")) {
                    if (!firstCell) builder.Append(" & ");
                    AppendLatex(builder, cell);
                    firstCell = false;
                }
                firstRow = false;
            }
            builder.Append("\\end{matrix}");
        }

        private static void AppendEquationArrayLatex(StringBuilder builder, OpenXmlElement element) {
            builder.Append("\\begin{aligned}");
            bool first = true;
            foreach (OpenXmlElement expression in FindChildren(element, "e")) {
                if (!first) builder.Append(" \\\\ ");
                AppendLatex(builder, expression);
                first = false;
            }
            builder.Append("\\end{aligned}");
        }

        private static string ToLatexDelimiter(string value) {
            if (value.Length == 0) return ".";
            return value switch {
                "{" => "\\{",
                "}" => "\\}",
                "[" => "[",
                "]" => "]",
                "|" => "|",
                "\u2016" => "\\|",
                _ => EscapeLatex(value)
            };
        }

        private static string EscapeLatex(string? value) {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            var builder = new StringBuilder(value!.Length);
            foreach (char character in value) {
                switch (character) {
                    case '\\': builder.Append("\\backslash{}"); break;
                    case '{': builder.Append("\\{"); break;
                    case '}': builder.Append("\\}"); break;
                    case '#': builder.Append("\\#"); break;
                    case '$': builder.Append("\\$"); break;
                    case '%': builder.Append("\\%"); break;
                    case '&': builder.Append("\\&"); break;
                    case '_': builder.Append("\\_"); break;
                    case '^': builder.Append("\\hat{}"); break;
                    case '~': builder.Append("\\tilde{}"); break;
                    default: builder.Append(character); break;
                }
            }
            return builder.ToString();
        }

    }
}
