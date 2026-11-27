using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    /// <summary>
    /// Projects OMML equations for the native DOC writer: the linear text Word shows as a field result and the EQ field
    /// instruction that stands for the equation.
    /// </summary>
    internal static class LegacyDocMath {
        private const int MaximumProjectionDepth = 256;
        private const int M = OfficeXmlNamespaces.Math;

        internal static string GetText(LegacyDocElement element) {
            EnsureMaximumProjectionDepth(element);
            var builder = new StringBuilder();
            AppendText(builder, element);
            return builder.ToString();
        }

        internal static string ToEquationFieldInstruction(LegacyDocElement element) {
            var builder = new StringBuilder(" EQ ");
            AppendEquationField(builder, element);
            builder.Append(" \\* MERGEFORMAT ");
            return builder.ToString();
        }

        private static void AppendText(StringBuilder builder, LegacyDocElement element) {
            if (element.Is(M, "t")) {
                builder.Append(element.Text);
                return;
            }

            // Wrappers can contain OMML, but similarly named drawing elements
            // such as VML's geometry formula v:f are not math fractions.
            if (element.Namespace != M) {
                foreach (LegacyDocElement child in element.Elements()) AppendText(builder, child);
                return;
            }

            switch (element.LocalName) {
                case "f":
                    AppendFractionText(builder, element);
                    return;
                case "sSup":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "^", element, "sup");
                    return;
                case "sSub":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "_", element, "sub");
                    return;
                case "sSubSup":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "_", element, "sub");
                    AppendScriptText(builder, "^", element, "sup");
                    return;
                case "sPre":
                    AppendScriptText(builder, "^", element, "sup");
                    AppendScriptText(builder, "_", element, "sub");
                    AppendChildText(builder, element, "e");
                    return;
                case "rad":
                    int degreeStart = builder.Length;
                    AppendChildText(builder, element, "deg");
                    if (builder.Length == degreeStart) {
                        builder.Append("sqrt(");
                    } else {
                        builder.Insert(degreeStart, "root(");
                        builder.Append(',');
                    }
                    AppendChildText(builder, element, "e");
                    builder.Append(')');
                    return;
                case "nary":
                case "int":
                    AppendNaryText(builder, element);
                    return;
                case "func":
                    int functionStart = builder.Length;
                    AppendChildText(builder, element, "fName");
                    if (builder.Length > functionStart) {
                        builder.Append('(');
                        AppendChildText(builder, element, "e");
                        builder.Append(')');
                    } else {
                        AppendChildText(builder, element, "e");
                    }
                    return;
                case "acc":
                    AppendAccentText(builder, element);
                    return;
                case "bar":
                    AppendFunctionText(builder, "bar", element, "e");
                    return;
                case "d":
                    AppendDelimiterText(builder, element);
                    return;
                case "groupChr":
                    AppendGroupCharacterText(builder, element);
                    return;
                case "m":
                    AppendMatrixText(builder, element);
                    return;
                case "eqArr":
                    builder.Append("eqarray(");
                    AppendJoinedChildText(builder, element, "e", ";");
                    builder.Append(')');
                    return;
                case "limLow":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "_", element, "lim");
                    return;
                case "limUpp":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "^", element, "lim");
                    return;
            }

            foreach (LegacyDocElement child in element.Elements()) {
                AppendText(builder, child);
            }
        }

        private static void AppendFractionText(StringBuilder builder, LegacyDocElement element) {
            switch (ReadFractionType(element)) {
                case MathFractionType.Linear:
                    AppendChildText(builder, element, "num");
                    builder.Append('/');
                    AppendChildText(builder, element, "den");
                    return;
                case MathFractionType.NoBar:
                    builder.Append("stack(");
                    AppendChildText(builder, element, "num");
                    builder.Append(',');
                    AppendChildText(builder, element, "den");
                    builder.Append(')');
                    return;
                case MathFractionType.Skewed:
                    AppendChildText(builder, element, "num");
                    builder.Append('⁄');
                    AppendChildText(builder, element, "den");
                    return;
                default:
                    builder.Append('(');
                    AppendChildText(builder, element, "num");
                    builder.Append(")/(");
                    AppendChildText(builder, element, "den");
                    builder.Append(')');
                    return;
            }
        }

        private static void AppendAccentText(StringBuilder builder, LegacyDocElement element) {
            string accent = ReadCharacterOrDefault(element, "chr", "̂");
            string functionName = accent switch {
                "^" => "hat",
                "̂" => "hat",
                "~" => "tilde",
                "̃" => "tilde",
                "." => "dot",
                "̇" => "dot",
                "¨" => "ddot",
                "̈" => "ddot",
                _ => string.Empty
            };
            if (functionName.Length > 0) {
                AppendFunctionText(builder, functionName, element, "e");
            } else {
                builder.Append("accent(");
                builder.Append(accent);
                builder.Append(',');
                AppendChildText(builder, element, "e");
                builder.Append(')');
            }
        }

        private static void AppendDelimiterText(StringBuilder builder, LegacyDocElement element) {
            MathCharacter begin = ReadCharacter(element, "begChr");
            MathCharacter end = ReadCharacter(element, "endChr");
            builder.Append(begin.Present ? begin.Value : "(");
            AppendJoinedChildText(builder, element, "e", ReadDelimiterSeparator(element));
            builder.Append(end.Present ? end.Value : ")");
        }

        private static void AppendGroupCharacterText(StringBuilder builder, LegacyDocElement element) {
            string character = ReadCharacterOrDefault(element, "chr", "⏟");
            string functionName = character switch {
                "⏞" => "overbrace",
                "⏟" => "underbrace",
                "⎴" => "overbracket",
                "⎵" => "underbracket",
                _ => "group"
            };
            AppendFunctionText(builder, functionName, element, "e");
        }

        private static void AppendMatrixText(StringBuilder builder, LegacyDocElement element) {
            builder.Append("matrix(");
            bool firstRow = true;
            foreach (LegacyDocElement row in element.Elements(M, "mr")) {
                if (!firstRow) builder.Append(';');
                bool firstCell = true;
                foreach (LegacyDocElement cell in row.Elements(M, "e")) {
                    if (!firstCell) builder.Append(',');
                    AppendText(builder, cell);
                    firstCell = false;
                }
                firstRow = false;
            }
            builder.Append(')');
        }

        private static void AppendFunctionText(StringBuilder builder, string functionName, LegacyDocElement element, string expressionLocalName) {
            builder.Append(functionName);
            builder.Append('(');
            AppendChildText(builder, element, expressionLocalName);
            builder.Append(')');
        }

        private static void AppendNaryText(StringBuilder builder, LegacyDocElement element) {
            builder.Append(ReadNaryOperatorText(element));
            AppendScriptText(builder, "_", element, "sub");
            AppendScriptText(builder, "^", element, "sup");
            int expressionStart = builder.Length;
            AppendChildText(builder, element, "e");
            if (builder.Length > expressionStart) {
                builder.Insert(expressionStart, "(");
                builder.Append(')');
            }
        }

        private static void AppendJoinedChildText(StringBuilder builder, LegacyDocElement element, string localName, string separator) {
            bool first = true;
            foreach (LegacyDocElement child in element.Elements(M, localName)) {
                if (!first) builder.Append(separator);
                AppendText(builder, child);
                first = false;
            }
        }

        private static void AppendScriptText(StringBuilder builder, string marker, LegacyDocElement element, string valueLocalName) {
            int valueStart = builder.Length;
            AppendChildText(builder, element, valueLocalName);
            if (builder.Length == valueStart) return;
            builder.Insert(valueStart, marker + "(");
            builder.Append(')');
        }

        private static void AppendChildText(StringBuilder builder, LegacyDocElement element, string localName) {
            LegacyDocElement child = element.Child(M, localName);
            if (child.Exists) AppendText(builder, child);
        }

        private static void AppendEquationField(StringBuilder builder, LegacyDocElement element) {
            if (element.Is(M, "t")) {
                builder.Append(EscapeEquationFieldLiteral(element.Text));
                return;
            }

            switch (element.LocalName) {
                case "f":
                    AppendFractionEquationField(builder, element);
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
                    builder.Append("\\)");
                    return;
                case "acc":
                    builder.Append("\\o(");
                    AppendEquationFieldChild(builder, element, "e");
                    builder.Append(',');
                    builder.Append(EscapeEquationFieldLiteral(ReadCharacterOrDefault(element, "chr", "̂")));
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

            foreach (LegacyDocElement child in element.Elements()) {
                AppendEquationField(builder, child);
            }
        }

        private static void AppendEquationFieldChild(StringBuilder builder, LegacyDocElement element, string localName) {
            LegacyDocElement child = element.Child(M, localName);
            if (child.Exists) AppendEquationField(builder, child);
        }

        private static void AppendFractionEquationField(StringBuilder builder, LegacyDocElement element) {
            switch (ReadFractionType(element)) {
                case MathFractionType.Linear:
                    AppendEquationFieldChild(builder, element, "num");
                    builder.Append('/');
                    AppendEquationFieldChild(builder, element, "den");
                    return;
                case MathFractionType.NoBar:
                    builder.Append("\\a\\co1(");
                    AppendEquationFieldChild(builder, element, "num");
                    builder.Append(',');
                    AppendEquationFieldChild(builder, element, "den");
                    builder.Append(')');
                    return;
                case MathFractionType.Skewed:
                    AppendEquationFieldChild(builder, element, "num");
                    builder.Append('⁄');
                    AppendEquationFieldChild(builder, element, "den");
                    return;
                default:
                    builder.Append("\\f(");
                    AppendEquationFieldChild(builder, element, "num");
                    builder.Append(',');
                    AppendEquationFieldChild(builder, element, "den");
                    builder.Append(')');
                    return;
            }
        }

        private static void AppendEquationFieldScript(StringBuilder builder, string switchName, LegacyDocElement element, string localName) {
            LegacyDocElement child = element.Child(M, localName);
            if (child.IsNone || GetText(child).Length == 0) return;
            builder.Append("\\s");
            builder.Append(switchName);
            builder.Append('(');
            AppendEquationField(builder, child);
            builder.Append(')');
        }

        private static void AppendNaryEquationField(StringBuilder builder, LegacyDocElement element) {
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

        private static void AppendDelimiterEquationField(StringBuilder builder, LegacyDocElement element) {
            MathCharacter begin = ReadCharacter(element, "begChr");
            MathCharacter end = ReadCharacter(element, "endChr");
            string beginValue = begin.Present ? begin.Value : "(";
            string endValue = end.Present ? end.Value : ")";
            string separator = EscapeEquationFieldLiteral(ReadDelimiterSeparator(element));
            if (beginValue.Length == 0 || endValue.Length == 0) {
                builder.Append(EscapeEquationFieldLiteral(beginValue));
                AppendJoinedEquationFieldChildren(builder, element, "e", separator);
                builder.Append(EscapeEquationFieldLiteral(endValue));
                return;
            }
            builder.Append("\\b\\lc");
            builder.Append(EscapeEquationFieldDelimiter(beginValue));
            builder.Append("\\rc");
            builder.Append(EscapeEquationFieldDelimiter(endValue));
            builder.Append('(');
            AppendJoinedEquationFieldChildren(builder, element, "e", separator);
            builder.Append(')');
        }

        private static void AppendMatrixEquationField(StringBuilder builder, LegacyDocElement element) {
            List<LegacyDocElement> rows = element.Elements(M, "mr").ToList();
            int columns = rows.Count == 0 ? 1 : Math.Max(1, rows[0].Elements(M, "e").Count);
            builder.Append("\\a\\co").Append(columns).Append('(');
            bool first = true;
            foreach (LegacyDocElement row in rows) {
                foreach (LegacyDocElement cell in row.Elements(M, "e")) {
                    if (!first) builder.Append(',');
                    AppendEquationField(builder, cell);
                    first = false;
                }
            }
            builder.Append(')');
        }

        private static void AppendEquationArrayField(StringBuilder builder, LegacyDocElement element) {
            builder.Append("\\a\\co1(");
            AppendJoinedEquationFieldChildren(builder, element, "e", ",");
            builder.Append(')');
        }

        private static void AppendJoinedEquationFieldChildren(StringBuilder builder, LegacyDocElement element, string localName, string separator) {
            bool first = true;
            foreach (LegacyDocElement child in element.Elements(M, localName)) {
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

        private static void EnsureMaximumProjectionDepth(LegacyDocElement element) {
            var pending = new Stack<(LegacyDocElement Element, int Depth)>();
            pending.Push((element, 1));
            while (pending.Count > 0) {
                (LegacyDocElement current, int depth) = pending.Pop();
                if (depth > MaximumProjectionDepth) {
                    throw new InvalidDataException(
                        "OMML equation nesting exceeds the configured " + MaximumProjectionDepth + "-level projection limit.");
                }
                foreach (LegacyDocElement child in current.Elements()) pending.Push((child, depth + 1));
            }
        }

        private static string ReadNaryOperatorText(LegacyDocElement element) {
            if (element.LocalName == "int") return "int";
            MathCharacter character = ReadCharacter(element, "chr");
            if (!character.Present) return "sum";
            return character.Value switch {
                "∑" => "sum",
                "∏" => "prod",
                "∫" => "int",
                _ => character.Value
            };
        }

        private static string ReadCharacterOrDefault(LegacyDocElement element, string localName, string defaultValue) {
            MathCharacter character = ReadCharacter(element, localName);
            return character.Present ? character.Value : defaultValue;
        }

        private static string ReadDelimiterSeparator(LegacyDocElement element) =>
            ReadCharacterOrDefault(element, "sepChr", "│");

        private static MathFractionType ReadFractionType(LegacyDocElement element) {
            LegacyDocElement type = element.Child(M, "fPr").Child(M, "type");
            string? value = ReadMathValue(type);
            return value switch {
                "lin" => MathFractionType.Linear,
                "noBar" => MathFractionType.NoBar,
                "skw" => MathFractionType.Skewed,
                _ => MathFractionType.Bar
            };
        }

        // The m:val attribute, or a val attribute without a namespace, as the SDK path accepted either.
        private static string? ReadMathValue(LegacyDocElement element) {
            foreach (LegacyDocAttribute attribute in element.Attributes) {
                if (attribute.LocalName == "val" && (attribute.Namespace == M || attribute.Namespace == OfficeXmlNamespaces.None)) {
                    return attribute.Value;
                }
            }
            return null;
        }

        private static MathCharacter ReadCharacter(LegacyDocElement element, string localName) {
            LegacyDocElement propertyContainer = LegacyDocElement.None;
            foreach (LegacyDocElement child in element.Elements()) {
                if (child.Namespace == M && child.LocalName.EndsWith("Pr", StringComparison.Ordinal)) {
                    propertyContainer = child;
                    break;
                }
            }
            LegacyDocElement character = propertyContainer.Child(M, localName);
            if (character.IsNone) return new MathCharacter(false, string.Empty);
            return new MathCharacter(true, ReadMathValue(character) ?? string.Empty);
        }

        private readonly struct MathCharacter {
            internal MathCharacter(bool present, string value) {
                Present = present;
                Value = value;
            }

            internal bool Present { get; }
            internal string Value { get; }
        }

        private enum MathFractionType {
            Bar,
            Linear,
            NoBar,
            Skewed
        }
    }
}
