using DocumentFormat.OpenXml;
using System.Text;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Word {
    /// <summary>
    /// Shared projections for Word mathematical content. OMML remains the editable source of truth;
    /// the other representations are deterministic fallbacks for text, legacy fields, and converters.
    /// </summary>
    internal static partial class WordMath {
        internal const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

        internal static string GetText(OpenXmlElement element) {
            var builder = new StringBuilder();
            AppendText(builder, element);
            return builder.ToString();
        }

        internal static void SetText(OpenXmlElement element, string? value) {
            string normalized = (value ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");

            if (element is M.Paragraph mathParagraph) {
                mathParagraph.RemoveAllChildren();
                mathParagraph.Append(new M.OfficeMath(new M.Run(new M.Text(normalized))));
                return;
            }

            if (element is OpenXmlCompositeElement composite) {
                composite.RemoveAllChildren();
                composite.Append(new M.Run(new M.Text(normalized)));
            }
        }

        private static void AppendText(StringBuilder builder, OpenXmlElement element) {
            if (element is M.Text text) {
                builder.Append(text.Text);
                return;
            }

            switch (element.LocalName) {
                case "f":
                    AppendFractionText(builder, element);
                    return;
                case "sSup":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "^", ReadChildText(element, "sup"));
                    return;
                case "sSub":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "_", ReadChildText(element, "sub"));
                    return;
                case "sSubSup":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "_", ReadChildText(element, "sub"));
                    AppendScriptText(builder, "^", ReadChildText(element, "sup"));
                    return;
                case "sPre":
                    AppendScriptText(builder, "^", ReadChildText(element, "sup"));
                    AppendScriptText(builder, "_", ReadChildText(element, "sub"));
                    AppendChildText(builder, element, "e");
                    return;
                case "rad":
                    string degree = ReadChildText(element, "deg");
                    string radicand = ReadChildText(element, "e");
                    builder.Append(degree.Length == 0 ? "sqrt(" : "root(" + degree + ",");
                    builder.Append(radicand);
                    builder.Append(')');
                    return;
                case "nary":
                case "int":
                    AppendNaryText(builder, element);
                    return;
                case "func":
                    string functionName = ReadChildText(element, "fName");
                    string argument = ReadChildText(element, "e");
                    if (functionName.Length > 0) {
                        AppendFunctionText(builder, functionName, argument);
                    } else {
                        builder.Append(argument);
                    }
                    return;
                case "acc":
                    AppendAccentText(builder, element);
                    return;
                case "bar":
                    AppendFunctionText(builder, "bar", ReadChildText(element, "e"));
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
                    AppendScriptText(builder, "_", ReadChildText(element, "lim"));
                    return;
                case "limUpp":
                    AppendChildText(builder, element, "e");
                    AppendScriptText(builder, "^", ReadChildText(element, "lim"));
                    return;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendText(builder, child);
            }
        }

        private static void AppendFractionText(StringBuilder builder, OpenXmlElement element) {
            string numerator = ReadChildText(element, "num");
            string denominator = ReadChildText(element, "den");
            switch (ReadFractionType(element)) {
                case MathFractionType.Linear:
                    builder.Append(numerator).Append('/').Append(denominator);
                    return;
                case MathFractionType.NoBar:
                    builder.Append("stack(").Append(numerator).Append(',').Append(denominator).Append(')');
                    return;
                case MathFractionType.Skewed:
                    builder.Append(numerator).Append('\u2044').Append(denominator);
                    return;
                default:
                    builder.Append('(').Append(numerator).Append(")/(").Append(denominator).Append(')');
                    return;
            }
        }

        private static void AppendAccentText(StringBuilder builder, OpenXmlElement element) {
            string expression = ReadChildText(element, "e");
            string accent = ReadCharacterOrDefault(element, "chr", "\u0302");
            string functionName = accent switch {
                "^" => "hat",
                "\u0302" => "hat",
                "~" => "tilde",
                "\u0303" => "tilde",
                "." => "dot",
                "\u0307" => "dot",
                "\u00a8" => "ddot",
                "\u0308" => "ddot",
                _ => string.Empty
            };
            if (functionName.Length > 0) {
                AppendFunctionText(builder, functionName, expression);
            } else {
                builder.Append("accent(");
                builder.Append(accent);
                builder.Append(',');
                builder.Append(expression);
                builder.Append(')');
            }
        }

        private static void AppendDelimiterText(StringBuilder builder, OpenXmlElement element) {
            MathCharacter begin = ReadCharacter(element, "begChr");
            MathCharacter end = ReadCharacter(element, "endChr");
            builder.Append(begin.Present ? begin.Value : "(");
            AppendJoinedChildText(builder, element, "e", ReadDelimiterSeparator(element));
            builder.Append(end.Present ? end.Value : ")");
        }

        private static void AppendGroupCharacterText(StringBuilder builder, OpenXmlElement element) {
            string character = ReadCharacterOrDefault(element, "chr", "\u23df");
            string functionName = character switch {
                "\u23de" => "overbrace",
                "\u23df" => "underbrace",
                "\u23b4" => "overbracket",
                "\u23b5" => "underbracket",
                _ => "group"
            };
            AppendFunctionText(builder, functionName, ReadChildText(element, "e"));
        }

        private static void AppendMatrixText(StringBuilder builder, OpenXmlElement element) {
            builder.Append("matrix(");
            bool firstRow = true;
            foreach (OpenXmlElement row in FindChildren(element, "mr")) {
                if (!firstRow) builder.Append(';');
                bool firstCell = true;
                foreach (OpenXmlElement cell in FindChildren(row, "e")) {
                    if (!firstCell) builder.Append(',');
                    AppendText(builder, cell);
                    firstCell = false;
                }
                firstRow = false;
            }
            builder.Append(')');
        }

        private static void AppendFunctionText(StringBuilder builder, string functionName, string expression) {
            builder.Append(functionName);
            builder.Append('(');
            builder.Append(expression);
            builder.Append(')');
        }

        private static void AppendNaryText(StringBuilder builder, OpenXmlElement element) {
            builder.Append(ReadNaryOperatorText(element));
            AppendScriptText(builder, "_", ReadChildText(element, "sub"));
            AppendScriptText(builder, "^", ReadChildText(element, "sup"));
            string expression = ReadChildText(element, "e");
            if (expression.Length > 0) {
                builder.Append('(');
                builder.Append(expression);
                builder.Append(')');
            }
        }

        private static void AppendJoinedChildText(StringBuilder builder, OpenXmlElement element, string localName, string separator) {
            bool first = true;
            foreach (OpenXmlElement child in FindChildren(element, localName)) {
                if (!first) builder.Append(separator);
                AppendText(builder, child);
                first = false;
            }
        }

        private static void AppendScriptText(StringBuilder builder, string marker, string value) {
            if (value.Length == 0) return;
            builder.Append(marker);
            builder.Append('(');
            builder.Append(value);
            builder.Append(')');
        }

        private static void AppendChildText(StringBuilder builder, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            if (child != null) AppendText(builder, child);
        }

        private static string ReadChildText(OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            return child == null ? string.Empty : GetText(child);
        }

        private static string ReadNaryOperatorText(OpenXmlElement element) {
            if (element.LocalName == "int") return "int";
            MathCharacter character = ReadCharacter(element, "chr");
            if (!character.Present) return "sum";
            return character.Value switch {
                "\u2211" => "sum",
                "\u220F" => "prod",
                "\u222B" => "int",
                _ => character.Value
            };
        }

        private static string ReadCharacterOrDefault(OpenXmlElement element, string localName, string defaultValue) {
            MathCharacter character = ReadCharacter(element, localName);
            return character.Present ? character.Value : defaultValue;
        }

        private static string ReadDelimiterSeparator(OpenXmlElement element) =>
            ReadCharacterOrDefault(element, "sepChr", "\u2502");

        private static MathFractionType ReadFractionType(OpenXmlElement element) {
            OpenXmlElement? properties = FindFirstChild(element, "fPr");
            OpenXmlElement? type = properties == null ? null : FindFirstChild(properties, "type");
            string? value = type?.GetAttributes()
                .FirstOrDefault(attribute => attribute.LocalName == "val" &&
                    (attribute.NamespaceUri == MathNamespace || attribute.NamespaceUri.Length == 0))
                .Value;
            return value switch {
                "lin" => MathFractionType.Linear,
                "noBar" => MathFractionType.NoBar,
                "skw" => MathFractionType.Skewed,
                _ => MathFractionType.Bar
            };
        }

        private static MathCharacter ReadCharacter(OpenXmlElement element, string localName) {
            OpenXmlElement? propertyContainer = element.ChildElements.FirstOrDefault(child =>
                child.NamespaceUri == MathNamespace && child.LocalName.EndsWith("Pr", StringComparison.Ordinal));
            OpenXmlElement? character = propertyContainer?.ChildElements.FirstOrDefault(child =>
                child.NamespaceUri == MathNamespace && child.LocalName == localName);
            if (character == null) return new MathCharacter(false, string.Empty);
            foreach (OpenXmlAttribute attribute in character.GetAttributes()) {
                if (attribute.LocalName == "val" && (attribute.NamespaceUri == MathNamespace || attribute.NamespaceUri.Length == 0)) {
                    return new MathCharacter(true, attribute.Value ?? string.Empty);
                }
            }
            return new MathCharacter(true, string.Empty);
        }

        private static IEnumerable<OpenXmlElement> FindChildren(OpenXmlElement element, string localName) {
            foreach (OpenXmlElement child in element.ChildElements) {
                if (child.NamespaceUri == MathNamespace && child.LocalName == localName) yield return child;
            }
        }

        private static OpenXmlElement? FindFirstChild(OpenXmlElement element, string localName) =>
            FindChildren(element, localName).FirstOrDefault();

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
