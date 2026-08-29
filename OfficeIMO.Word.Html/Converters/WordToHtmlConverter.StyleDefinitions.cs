using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private WordDocument? _styleDecorationDocument;
        private IReadOnlyDictionary<string, Style>? _styleDecorationDefinitions;
        private readonly Dictionary<string, WordStyleTextDecorations> _styleDecorationCache =
            new(StringComparer.OrdinalIgnoreCase);

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

            var stylePart = document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart;
            var styleMap = (stylePart?.Styles?.OfType<Style>() ?? Enumerable.Empty<Style>())
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .GroupBy(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            string BuildCss(string styleId, bool includeInlineVerticalAlignment) {
                var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                void Merge(string id) {
                    var key = id;
                    if (string.IsNullOrEmpty(key)) {
                        return;
                    }
                    if (!visited.Add(key)) {
                        return;
                    }
                    if (!styleMap.TryGetValue(key, out var def)) {
                        return;
                    }
                    var baseId = def.BasedOn?.Val;
                    if (!string.IsNullOrEmpty(baseId)) {
                        Merge(baseId!);
                    }
                    var pPr = def.StyleParagraphProperties;
                    if (pPr?.Justification?.Val != null) {
                        var justifyVal = pPr.Justification.Val.Value;
                        var alignment = "left";
                        if (justifyVal == JustificationValues.Center) {
                            alignment = "center";
                        } else if (justifyVal == JustificationValues.Right) {
                            alignment = "right";
                        } else if (justifyVal == JustificationValues.Both) {
                            alignment = "justify";
                        }
                        props["text-align"] = alignment;
                    }
                    var rPr = def.StyleRunProperties;
                    if (rPr != null) {
                        if (rPr.Bold != null) {
                            props["font-weight"] = "bold";
                        }
                        if (rPr.Italic != null) {
                            props["font-style"] = "italic";
                        }
                        // Decorations are emitted as independent nested wrappers at each style
                        // use site. CSS exposes one text-decoration-style value, so a single rule
                        // cannot faithfully combine patterns such as wavy underline + double strike.
                        if (includeInlineVerticalAlignment) {
                            if (rPr.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Superscript) {
                                props["vertical-align"] = "super";
                            } else if (rPr.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Subscript) {
                                props["vertical-align"] = "sub";
                            } else if (rPr.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Baseline) {
                                props["vertical-align"] = "baseline";
                            }
                        }
                        if (IsEnabled(rPr.SmallCaps)) {
                            props["font-variant"] = "small-caps";
                        } else if (IsEnabled(rPr.Caps)) {
                            props["text-transform"] = "uppercase";
                        }
                        var colorVal = rPr.Color?.Val?.Value;
                        if (!string.IsNullOrEmpty(colorVal) &&
                            !string.Equals(colorVal, "auto", StringComparison.OrdinalIgnoreCase) &&
                            IsSixDigitHexColor(colorVal!)) {
                            var cv = colorVal!;
                            props["color"] = "#" + cv.ToLowerInvariant();
                        }
                        var sizeVal = rPr.FontSize?.Val;
                        if (!string.IsNullOrEmpty(sizeVal) && int.TryParse(sizeVal, out int sz)) {
                            props["font-size"] = (sz / 2.0).ToString("0.##") + "pt";
                        }
                        var font = rPr.RunFonts?.Ascii?.Value;
                        if (!string.IsNullOrEmpty(font)) {
                            props["font-family"] = QuoteCssString(font!);
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
                _styleDecorationDefinitions = document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles?
                    .OfType<Style>()
                    .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                    .GroupBy(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            }
            if (_styleDecorationCache.TryGetValue(styleId!, out WordStyleTextDecorations? cached)) return cached;
            IReadOnlyDictionary<string, Style>? styleMap = _styleDecorationDefinitions;
            if (styleMap == null) return new WordStyleTextDecorations();
            var chain = new Stack<Style>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string? current = styleId;
            while (!string.IsNullOrWhiteSpace(current) && visited.Add(current!) && styleMap.TryGetValue(current!, out Style? style)) {
                chain.Push(style);
                current = style.BasedOn?.Val?.Value;
            }

            var result = new WordStyleTextDecorations();
            while (chain.Count > 0) {
                StyleRunProperties? properties = chain.Pop().StyleRunProperties;
                if (properties?.Underline?.Val?.Value is UnderlineValues underline) {
                    result.Underline = underline == UnderlineValues.None ? WordUnderlineStyle.None : underline.ToOfficeEnum();
                }
                if (properties?.Strike != null) result.Strike = IsEnabled(properties.Strike);
                if (properties?.DoubleStrike != null) result.DoubleStrike = IsEnabled(properties.DoubleStrike);
                if (properties?.VerticalTextAlignment?.Val?.Value is VerticalPositionValues verticalPosition) {
                    result.VerticalPosition = verticalPosition.ToOfficeEnum();
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
