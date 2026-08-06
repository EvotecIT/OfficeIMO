using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Word {
    /// <summary>
    /// Shared projections for Word mathematical content. OMML remains the editable source of truth;
    /// the other representations are deterministic fallbacks for text, legacy fields, and converters.
    /// </summary>
    internal static partial class WordMath {
        internal const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        internal const int DefaultMaximumProjectionDepth = 256;

        internal static string GetText(OpenXmlElement element) =>
            GetText(element, DefaultMaximumProjectionDepth);

        internal static string GetText(OpenXmlElement element, int maxDepth) {
            return GetText(element, long.MaxValue, maxDepth);
        }

        internal static string GetText(OpenXmlElement element, long maxCharacters, int maxDepth) {
            EnsureMaximumProjectionDepth(element, maxDepth);
            return GetTextValidated(element, maxCharacters);
        }

        private static string GetTextValidated(OpenXmlElement element, long maxCharacters = long.MaxValue) {
            var builder = new BoundedStringBuilder(maxCharacters);
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

        private static void AppendText(BoundedStringBuilder builder, OpenXmlElement element) {
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

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendText(builder, child);
            }
        }

        private static void AppendFractionText(BoundedStringBuilder builder, OpenXmlElement element) {
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
                    builder.Append('\u2044');
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

        private static void AppendAccentText(BoundedStringBuilder builder, OpenXmlElement element) {
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
                AppendFunctionText(builder, functionName, element, "e");
            } else {
                builder.Append("accent(");
                builder.Append(accent);
                builder.Append(',');
                AppendChildText(builder, element, "e");
                builder.Append(')');
            }
        }

        private static void AppendDelimiterText(BoundedStringBuilder builder, OpenXmlElement element) {
            MathCharacter begin = ReadCharacter(element, "begChr");
            MathCharacter end = ReadCharacter(element, "endChr");
            builder.Append(begin.Present ? begin.Value : "(");
            AppendJoinedChildText(builder, element, "e", ReadDelimiterSeparator(element));
            builder.Append(end.Present ? end.Value : ")");
        }

        private static void AppendGroupCharacterText(BoundedStringBuilder builder, OpenXmlElement element) {
            string character = ReadCharacterOrDefault(element, "chr", "\u23df");
            string functionName = character switch {
                "\u23de" => "overbrace",
                "\u23df" => "underbrace",
                "\u23b4" => "overbracket",
                "\u23b5" => "underbracket",
                _ => "group"
            };
            AppendFunctionText(builder, functionName, element, "e");
        }

        private static void AppendMatrixText(BoundedStringBuilder builder, OpenXmlElement element) {
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

        private static void AppendFunctionText(
            BoundedStringBuilder builder,
            string functionName,
            OpenXmlElement element,
            string expressionLocalName) {
            builder.Append(functionName);
            builder.Append('(');
            AppendChildText(builder, element, expressionLocalName);
            builder.Append(')');
        }

        private static void AppendNaryText(BoundedStringBuilder builder, OpenXmlElement element) {
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

        private static void AppendJoinedChildText(BoundedStringBuilder builder, OpenXmlElement element, string localName, string separator) {
            bool first = true;
            foreach (OpenXmlElement child in FindChildren(element, localName)) {
                if (!first) builder.Append(separator);
                AppendText(builder, child);
                first = false;
            }
        }

        private static void AppendScriptText(
            BoundedStringBuilder builder,
            string marker,
            OpenXmlElement element,
            string valueLocalName) {
            int valueStart = builder.Length;
            AppendChildText(builder, element, valueLocalName);
            if (builder.Length == valueStart) return;
            builder.Insert(valueStart, marker + "(");
            builder.Append(')');
        }

        private static void AppendChildText(BoundedStringBuilder builder, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            if (child != null) AppendText(builder, child);
        }

        private static string ReadChildText(OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            return child == null ? string.Empty : GetTextValidated(child);
        }

        private static void EnsureMaximumProjectionDepth(OpenXmlElement element, int maxDepth) {
            if (maxDepth <= 0) throw new ArgumentOutOfRangeException(nameof(maxDepth));
            var pending = new Stack<(OpenXmlElement Element, int Depth)>();
            pending.Push((element, 1));
            while (pending.Count > 0) {
                (OpenXmlElement current, int depth) = pending.Pop();
                if (depth > maxDepth) {
                    throw new InvalidDataException(
                        "OMML equation nesting exceeds the configured " + maxDepth + "-level projection limit.");
                }
                for (int index = current.ChildElements.Count - 1; index >= 0; index--) {
                    pending.Push((current.ChildElements[index], depth + 1));
                }
            }
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
