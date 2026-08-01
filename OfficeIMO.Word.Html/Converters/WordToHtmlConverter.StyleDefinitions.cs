using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
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

            string BuildCss(string styleId) {
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
                        var underline = rPr.Underline?.Val?.Value;
                        if (underline != null && underline != UnderlineValues.None) {
                            props["text-decoration"] = "underline";
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
                var css = BuildCss(s);
                sb.Append('.').Append(GetSafeStyleClassName(s)).Append(" { ").Append(css).Append(" }\n");
            }
            foreach (var s in runStyles) {
                cancellationToken.ThrowIfCancellationRequested();
                var css = BuildCss(s);
                sb.Append('.').Append(GetSafeStyleClassName(s)).Append(" { ").Append(css).Append(" }\n");
            }
            styleElement.TextContent = sb.ToString();
            head.AppendChild(styleElement);
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
