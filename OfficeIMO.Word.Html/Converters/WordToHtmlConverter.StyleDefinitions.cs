using AngleSharp.Dom;
using System.Text;
using System.Security.Cryptography;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private WordDocument? _styleDecorationDocument;
        private WordStyleCatalog? _styleDecorationDefinitions;
        private readonly Dictionary<string, WordStyleTextDecorations> _styleDecorationCache =
            new(StringComparer.Ordinal);

        private static void AppendStyleDefinitions(
            WordDocument document,
            IDocument htmlDoc,
            IElement head,
            HashSet<string> paragraphStyles,
            HashSet<string> runStyles,
            CancellationToken cancellationToken) {
            if (paragraphStyles.Count == 0 && runStyles.Count == 0) {
                return;
            }

            WordStyleCatalog styleMap = WordStyleCatalog.For(document);

            string BuildCss(string styleId, bool includeInlineVerticalAlignment) {
                var visited = new HashSet<string>(StringComparer.Ordinal);
                var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                void Merge(string id) {
                    var key = id;
                    if (string.IsNullOrEmpty(key)) {
                        return;
                    }
                    if (!visited.Add(key)) {
                        return;
                    }
                    OfficeXmlNode def = styleMap.FindStyleNodeById(key);
                    if (def.IsNone) {
                        return;
                    }
                    var baseId = def.WChild("basedOn").WVal();
                    if (!string.IsNullOrEmpty(baseId)) {
                        Merge(baseId!);
                    }
                    var pPr = def.WChild("pPr");
                    if (pPr.WChild("jc").WVal() is { } justifyVal) {
                        var alignment = "left";
                        if (justifyVal == "center") {
                            alignment = "center";
                        } else if (justifyVal == "right") {
                            alignment = "right";
                        } else if (justifyVal == "both") {
                            alignment = "justify";
                        }
                        props["text-align"] = alignment;
                    }
                    var rPr = def.WChild("rPr");
                    if (rPr.Exists) {
                        if (rPr.WChild("b").Exists) {
                            props["font-weight"] = "bold";
                        }
                        if (rPr.WChild("i").Exists) {
                            props["font-style"] = "italic";
                        }
                        // Decorations are emitted as independent nested wrappers at each style
                        // use site. CSS exposes one text-decoration-style value, so a single rule
                        // cannot faithfully combine patterns such as wavy underline + double strike.
                        if (includeInlineVerticalAlignment) {
                            // Superscript and subscript are emitted as structural wrappers so their
                            // semantics survive without CSS. Only an explicit baseline needs a class
                            // declaration to reset a script inherited from a base style.
                            if (rPr.WChild("vertAlign").WVal() == "baseline") {
                                props["vertical-align"] = "baseline";
                            }
                        }
                        if (rPr.WChild("smallCaps").Exists) {
                            props["font-variant"] = IsEnabled(rPr.WChild("smallCaps")) ? "small-caps" : "normal";
                        }
                        if (rPr.WChild("caps").Exists) {
                            props["text-transform"] = IsEnabled(rPr.WChild("caps")) ? "uppercase" : "none";
                        }
                        var colorVal = WordConverterNodes.ResolveColorHex(rPr.WChild("color"), document);
                        if (!string.IsNullOrEmpty(colorVal) &&
                            !string.Equals(colorVal, "auto", StringComparison.OrdinalIgnoreCase) &&
                            IsSixDigitHexColor(colorVal!)) {
                            var cv = colorVal!;
                            props["color"] = "#" + cv.ToLowerInvariant();
                        }
                        var sizeVal = rPr.WChild("sz").WVal();
                        if (!string.IsNullOrEmpty(sizeVal) && int.TryParse(sizeVal, out int sz)) {
                            props["font-size"] = (sz / 2.0).ToString("0.##") + "pt";
                        }
                        var font = ResolveRunFonts(rPr.WChild("rFonts"), document);
                        if (!string.IsNullOrEmpty(font)) {
                            props["font-family"] = FontFamilyCss(font!);
                        }
                    }
                }

                Merge(styleId);

                return string.Join(" ", props.Select(kv => kv.Key + ':' + kv.Value + ';'));
            }

            var styleElement = CreateOutputElement(htmlDoc, "style");
            var sb = new StringBuilder();

            foreach (var s in paragraphStyles) {
                cancellationToken.ThrowIfCancellationRequested();
                var css = BuildCss(s, includeInlineVerticalAlignment: false);
                AppendCssRule(s, css);
            }
            foreach (var s in runStyles) {
                cancellationToken.ThrowIfCancellationRequested();
                var css = BuildCss(s, includeInlineVerticalAlignment: true);
                AppendCssRule(s, css);
            }
            styleElement.TextContent = sb.ToString();
            head.AppendChild(styleElement);

            void AppendCssRule(string styleId, string css) {
                string rule = "." + GetSafeStyleClassName(styleId) + " { " + css + " }\n";
                ReserveOutputCharacters(
                    htmlDoc,
                    rule.Length,
                    "Generated style CSS exceeds the configured output-character limit before DOM construction.",
                    "GeneratedStyleCss");
                sb.Append(rule);
            }
        }

        private INode ApplyStyleDefinitionTextDecorations(
            WordDocument document,
            IDocument htmlDocument,
            string? styleId,
            INode node,
            string source,
            bool suppressUnderline = false,
            bool suppressStrike = false,
            bool suppressDoubleStrike = false,
            bool suppressVerticalPosition = false) {
            WordStyleTextDecorations decorations = ResolveStyleDefinitionTextDecorations(document, styleId);
            if (decorations.DoubleStrike && !suppressDoubleStrike) {
                var strike = CreateOutputElement(htmlDocument, "span");
                SetOutputAttribute(htmlDocument, strike, "style", "text-decoration-line:line-through;text-decoration-style:double", source + ":double-strike");
                SetOutputAttribute(htmlDocument, strike, "data-officeimo-word-double-strike", "true", source + ":double-strike-metadata");
                strike.AppendChild(node);
                node = strike;
            } else if (decorations.Strike && !suppressStrike) {
                var strike = CreateOutputElement(htmlDocument, "s");
                strike.AppendChild(node);
                node = strike;
            }

            if (!suppressUnderline && decorations.Underline.HasValue && decorations.Underline.Value != WordUnderlineStyle.None) {
                if (decorations.Underline.Value == WordUnderlineStyle.Single) {
                    var underline = CreateOutputElement(htmlDocument, "u");
                    underline.AppendChild(node);
                    node = underline;
                } else {
                    var underline = CreateOutputElement(htmlDocument, "span");
                    SetOutputAttribute(htmlDocument, underline, "style",
                        "text-decoration-line:underline;text-decoration-style:" + MapWordUnderlineToCssStyle(decorations.Underline.Value),
                        source + ":underline");
                    SetOutputAttribute(htmlDocument, underline, "data-officeimo-word-underline", decorations.Underline.Value.ToString(),
                        source + ":underline-metadata");
                    underline.AppendChild(node);
                    node = underline;
                }
            }

            if (!suppressVerticalPosition) {
                if (decorations.VerticalPosition == WordVerticalTextPosition.Superscript) {
                    var superscript = CreateOutputElement(htmlDocument, "sup");
                    superscript.AppendChild(node);
                    node = superscript;
                } else if (decorations.VerticalPosition == WordVerticalTextPosition.Subscript) {
                    var subscript = CreateOutputElement(htmlDocument, "sub");
                    subscript.AppendChild(node);
                    node = subscript;
                }
            }

            return node;
        }

        private WordStyleTextDecorations ResolveStyleDefinitionTextDecorations(WordDocument document, string? styleId) {
            if (string.IsNullOrWhiteSpace(styleId)) return new WordStyleTextDecorations();
            if (!ReferenceEquals(_styleDecorationDocument, document)) {
                _styleDecorationDocument = document;
                _styleDecorationCache.Clear();
                _styleDecorationDefinitions = WordStyleCatalog.For(document);
            }
            if (_styleDecorationCache.TryGetValue(styleId!, out WordStyleTextDecorations? cached)) return cached;
            WordStyleCatalog? styleMap = _styleDecorationDefinitions;
            if (styleMap == null) return new WordStyleTextDecorations();
            var chain = new Stack<OfficeXmlNode>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? current = styleId;
            while (!string.IsNullOrWhiteSpace(current) && visited.Add(current!) && styleMap.FindStyleNodeById(current) is { Exists: true } style) {
                chain.Push(style);
                current = style.WChild("basedOn").WVal();
            }

            var result = new WordStyleTextDecorations();
            while (chain.Count > 0) {
                OfficeXmlNode properties = chain.Pop().WChild("rPr");
                if (properties.WChild("u").WVal() is { } underline) {
                    result.Underline = ParseUnderlineToken(underline);
                }
                if (properties.WChild("strike").Exists) result.Strike = IsEnabled(properties.WChild("strike"));
                if (properties.WChild("dstrike").Exists) result.DoubleStrike = IsEnabled(properties.WChild("dstrike"));
                if (properties.WChild("vertAlign").WVal() is { } verticalPosition) {
                    result.VerticalPosition = ParseVerticalPositionToken(verticalPosition);
                }
            }
            _styleDecorationCache[styleId!] = result;
            return result;
        }

        private sealed class WordStyleTextDecorations {
            internal WordUnderlineStyle? Underline { get; set; }
            internal bool Strike { get; set; }
            internal bool DoubleStrike { get; set; }
            internal WordVerticalTextPosition? VerticalPosition { get; set; }
        }

        private static bool IsSixDigitHexColor(string value) {
            if (value.Length != 6) return false;
            for (int i = 0; i < value.Length; i++) {
                char character = value[i];
                if (!((character >= '0' && character <= '9')
                    || (character >= 'A' && character <= 'F')
                    || (character >= 'a' && character <= 'f'))) {
                    return false;
                }
            }
            return true;
        }

        private static string? NormalizeSixDigitHexColor(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim().TrimStart('#');
            return IsSixDigitHexColor(normalized) ? normalized.ToLowerInvariant() : null;
        }

        private static string FontFamilyCss(string value) {
            string generic = global::OfficeIMO.Drawing.OfficeDocumentFontSubstitutions.ClassifyGenericFamily(value);
            return string.Equals(value.Trim(), generic, StringComparison.OrdinalIgnoreCase)
                ? generic : QuoteCssString(value) + ", " + generic;
        }

        private static string QuoteCssString(string value) {
            var escaped = new StringBuilder(value.Length + 2);
            escaped.Append('"');
            foreach (char character in value) {
                if ((character >= 'A' && character <= 'Z')
                    || (character >= 'a' && character <= 'z')
                    || (character >= '0' && character <= '9')
                    || character == ' '
                    || character == '-'
                    || character == '_') {
                    escaped.Append(character);
                } else {
                    escaped.Append('\\')
                        .Append(((int)character).ToString("x", System.Globalization.CultureInfo.InvariantCulture))
                        .Append(' ');
                }
            }
            return escaped.Append('"').ToString();
        }

        private static string GetSafeStyleClassName(string? styleId) {
            if (styleId != null
                && styleId.Length > 0
                && IsSafeStyleClassStart(styleId[0])
                && styleId.All(IsSafeStyleClassCharacter)) {
                return styleId;
            }

            using SHA256 sha256 = SHA256.Create();
            byte[] digest = sha256.ComputeHash(Encoding.UTF8.GetBytes(styleId ?? string.Empty));
            var suffix = new StringBuilder(24);
            for (int i = 0; i < 12; i++) {
                suffix.Append(digest[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
            }
            return "word-style-" + suffix;
        }

        private static bool IsSafeStyleClassStart(char character) =>
            (character >= 'A' && character <= 'Z')
            || (character >= 'a' && character <= 'z')
            || character == '_';

        private static bool IsSafeStyleClassCharacter(char character) =>
            IsSafeStyleClassStart(character)
            || (character >= '0' && character <= '9')
            || character == '-';
    }
}
